<#
.SYNOPSIS
    Массовая настройка WinRM на Windows-хостах через PsExec.
.DESCRIPTION
    1. Читает inventory.yaml и фильтрует Windows-хосты.
    2. Для каждого хоста:
       - проверяет state из output/<domain>/winrm/<host>.json — если OK, пропускает (если нет -Force)
       - TCP 5985 — если открыт → WinRM уже работает, status = OK
       - TCP 445  — если закрыт → status = SMB_FAIL (PsExec не доступен)
       - PsExec → Enable-PSRemoting + Basic + AllowUnencrypted + firewall
       - повторная проверка WinRM
    3. Результат для каждого хоста сохраняется в output/<domain>/winrm/<host>.json.

.PARAMETER DryRun
    Только показать план, без изменений.
.PARAMETER HostFilter
    Wildcard-фильтр по имени или IP.
.PARAMETER HostsFile
    Путь к файлу со списком хостов (по одному на строку). Приоритетнее HostFilter.
.PARAMETER ConfigPath
    Путь к config.yaml. По умолчанию — рядом со скриптом.
.PARAMETER Force
    Игнорировать сохранённый state — заново обрабатывать все хосты, даже если OK.

.EXAMPLE
    .\Deploy-WinRM.ps1 -DryRun
    .\Deploy-WinRM.ps1 -HostsFile .\tmp\retry-hosts-winrm.txt
    .\Deploy-WinRM.ps1 -Force
#>

param(
    [switch]$DryRun,
    [string]$HostFilter = "*",
    [string]$HostsFile,
    [string]$ConfigPath,
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Импорт общей библиотеки
. (Join-Path $ScriptDir "Common.ps1")

Assert-Admin -ScriptPath $MyInvocation.MyCommand.Definition -BoundParams $PSBoundParameters `
    -SkipIf { Test-WSManClientReady } -Reason "настройка локального WSMan-клиента"

$LogFile = New-LogFile -RootDir $ScriptDir -ScriptName "Deploy-WinRM"

# ============================================================
# PsExec remote script
# ============================================================

function Invoke-PsExecWinRM {
    param(
        [string]$PsExecPath,
        [string]$ComputerName,
        [string]$Username,
        [string]$Password
    )

    $remoteScript = @'
try {
    Enable-PSRemoting -Force -SkipNetworkProfileCheck -ErrorAction Stop
    Set-Item WSMan:\localhost\Service\Auth\Basic -Value $true -Force
    Set-Item WSMan:\localhost\Service\AllowUnencrypted -Value $true -Force
    Set-Item WSMan:\localhost\Client\TrustedHosts -Value '*' -Force
    netsh advfirewall firewall add rule name="WinRM HTTP" dir=in action=allow protocol=TCP localport=5985 2>$null
    Write-Host "WINRM_OK"
} catch {
    Write-Host "WINRM_ERROR: $_"
    exit 1
}
'@

    $psexecArgs = @(
        "\\$ComputerName",
        "-u", $Username,
        "-p", $Password,
        "-accepteula",
        "-nobanner",
        "-h",
        "powershell.exe", "-ExecutionPolicy", "Bypass", "-Command", $remoteScript
    )

    $proc = Start-Process -FilePath $PsExecPath `
        -ArgumentList $psexecArgs `
        -NoNewWindow -Wait -PassThru `
        -RedirectStandardOutput "$env:TEMP\psexec_out.txt" `
        -RedirectStandardError "$env:TEMP\psexec_err.txt"

    $stdout = if (Test-Path "$env:TEMP\psexec_out.txt") { Get-Content "$env:TEMP\psexec_out.txt" -Raw } else { "" }
    $stderr = if (Test-Path "$env:TEMP\psexec_err.txt") { Get-Content "$env:TEMP\psexec_err.txt" -Raw } else { "" }

    return @{ ExitCode = $proc.ExitCode; StdOut = $stdout; StdErr = $stderr }
}

function Test-WinRMService {
    param([string]$ComputerName, [int]$TimeoutSec = 5)
    if (-not (Test-TcpPort -ComputerName $ComputerName -Port 5985 -TimeoutSec $TimeoutSec)) {
        return $false
    }
    try {
        $null = Test-WSMan -ComputerName $ComputerName -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

# ============================================================
# MAIN
# ============================================================

Write-Log "=========================================="
Write-Log "=== Deploy-WinRM — Начало ==="
Write-Log "=========================================="
Write-Log "Лог: $LogFile"

# --- Конфиг ---
if (-not $ConfigPath) { $ConfigPath = Join-Path $ScriptDir "config.yaml" }
if (-not (Test-Path $ConfigPath)) {
    Write-Log "Конфиг не найден: $ConfigPath" "ERROR"
    exit 1
}
Write-Log "Конфиг: $ConfigPath"
$config = Parse-SimpleYaml -Path $ConfigPath

$psexecPath    = $config['psexec_path']
$inventoryPath = $config['inventory_path']
$username      = $config['credentials']['username']
$winrmTimeout  = if ($config['winrm_test_timeout']) { [int]$config['winrm_test_timeout'] } else { 5 }
$domain        = if ($config['domain']) { $config['domain'] } else { '_default' }

if (-not (Test-Path $psexecPath)) {
    Write-Log "PsExec не найден: $psexecPath" "ERROR"
    exit 1
}
if (-not (Test-Path $inventoryPath)) {
    Write-Log "Inventory не найден: $inventoryPath" "ERROR"
    exit 1
}

Write-Log "Domain: $domain"

# --- Пароль ---
$password = $null
if (-not $DryRun) {
    $cfgPassword = $config['credentials']['password']
    if ($cfgPassword -and $cfgPassword -ne 'PUT_PASSWORD_HERE') {
        Write-Log "Пароль: из конфига"
        $password = $cfgPassword
    } else {
        Write-Log "Запрос пароля для $username..."
        $cred = Get-Credential -UserName $username -Message "Пароль для PsExec ($username)"
        if (-not $cred) { Write-Log "Отменено пользователем" "ERROR"; exit 1 }
        $password = $cred.GetNetworkCredential().Password
    }
}

# --- Inventory ---
Write-Log "Чтение inventory: $inventoryPath"
$inventory = Parse-SimpleYaml -Path $inventoryPath
$allHosts = Get-WindowsHosts -Inventory $inventory
$filteredHosts = Apply-HostsFilter -AllHosts $allHosts -HostFilter $HostFilter -HostsFile $HostsFile

Write-Log "Всего Windows-хостов: $($allHosts.Count)"
if ($HostsFile) {
    Write-Log "После HostsFile '$HostsFile': $($filteredHosts.Count)"
} else {
    Write-Log "После фильтра '$HostFilter': $($filteredHosts.Count)"
}

if ($filteredHosts.Count -eq 0) {
    Write-Log "Нет хостов для обработки" "WARN"
    exit 0
}

# --- State machine ---
$stateDir = Get-StateDir -RootDir $ScriptDir -Domain $domain -Transport "winrm"
Write-Log "State directory: $stateDir"
if ($Force) { Write-Log "Force: state будет игнорироваться, все хосты переобрабатываются" "WARN" }

# --- Обработка ---
Write-Log ""
Write-Log "=== Обработка ==="
Write-Log ""

$stats = @{
    OK = 0; SKIPPED = 0; ALREADY_WORKING = 0
    OFFLINE = 0; SMB_FAIL = 0; PSEXEC_FAIL = 0; VERIFY_FAIL = 0
}

foreach ($h in $filteredHosts) {
    $displayName = "$($h.Name) ($($h.IP))"
    $statePath = Get-HostStatePath -StateDir $stateDir -HostName $h.Name
    $prevState = Read-HostState -Path $statePath

    if (Test-ShouldSkip -State $prevState -Force:$Force) {
        Write-Log "  [ SKIP ] $displayName — уже OK ($($prevState.timestamp))"
        $stats.SKIPPED++
        continue
    }

    $state = @{
        host = $h.Name; ip = $h.IP; transport = ''
        status = 'UNKNOWN'; category = ''; exit_code = $null; message = ''
    }

    # 1. WinRM уже работает?
    if (Test-WinRMService -ComputerName $h.IP -TimeoutSec $winrmTimeout) {
        Write-Log "  [ OK ] $displayName — WinRM уже работает" "OK"
        $state.status = 'OK'
        $state.category = 'ALREADY_WORKING'
        $state.transport = 'probe'   # никакой настройки не делали, только проверили
        $state.message = 'WinRM was already running'
        $stats.OK++
        $stats.ALREADY_WORKING++
        Write-HostState -Path $statePath -State $state
        continue
    }

    # 2. Хост вообще онлайн?
    $ping = Test-Connection -ComputerName $h.IP -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $ping) {
        Write-Log "  [ OFFLINE ] $displayName" "WARN"
        $state.status = 'OFFLINE'; $state.category = 'OFFLINE'
        $state.transport = 'probe'
        $state.message = 'ICMP not responding'
        $stats.OFFLINE++
        Write-HostState -Path $statePath -State $state
        continue
    }

    # 3. SMB доступен для PsExec?
    if (-not (Test-TcpPort -ComputerName $h.IP -Port 445 -TimeoutSec $winrmTimeout)) {
        Write-Log "  [ SMB_FAIL ] $displayName — TCP 445 закрыт, PsExec не сможет подключиться" "WARN"
        $state.status = 'SMB_FAIL'; $state.category = 'SMB_UNREACHABLE'
        $state.transport = 'probe'
        $state.message = 'TCP 445 closed — admin$/PsExec path unavailable'
        $stats.SMB_FAIL++
        Write-HostState -Path $statePath -State $state
        continue
    }

    if ($DryRun) {
        Write-Log "  [ DRY ] $displayName — будет настройка через PsExec"
        continue
    }

    # 4. PsExec → настройка WinRM
    Write-Log "  Настройка: $displayName"
    $state.transport = 'psexec'
    try {
        $r = Invoke-PsExecWinRM -PsExecPath $psexecPath -ComputerName $h.IP `
            -Username $username -Password $password
        $state.exit_code = $r.ExitCode
        if ($r.ExitCode -ne 0) {
            Write-Log "  [ PSEXEC_FAIL ] $displayName — exitcode=$($r.ExitCode): $($r.StdErr)" "ERROR"
            $state.status = 'PSEXEC_FAIL'; $state.category = 'PSEXEC_ERROR'
            $state.message = $r.StdErr.Trim()
            $stats.PSEXEC_FAIL++
            Write-HostState -Path $statePath -State $state
            continue
        }
    } catch {
        Write-Log "  [ PSEXEC_FAIL ] $displayName — exception: $($_.Exception.Message)" "ERROR"
        $state.status = 'PSEXEC_FAIL'; $state.category = 'EXCEPTION'
        $state.message = $_.Exception.Message
        $stats.PSEXEC_FAIL++
        Write-HostState -Path $statePath -State $state
        continue
    }

    # 5. Verify
    Start-Sleep -Seconds 2
    if (Test-WinRMService -ComputerName $h.IP -TimeoutSec $winrmTimeout) {
        Write-Log "  [ OK ] $displayName — WinRM настроен и работает" "OK"
        $state.status = 'OK'; $state.category = 'CONFIGURED'
        $state.message = 'PsExec configured WinRM, verified'
        $stats.OK++
        Write-HostState -Path $statePath -State $state
    } else {
        Write-Log "  [ VERIFY_FAIL ] $displayName — после настройки WinRM не отвечает" "ERROR"
        $state.status = 'VERIFY_FAIL'; $state.category = 'VERIFY_FAIL'
        $state.message = 'PsExec returned OK, but WinRM still unreachable'
        $stats.VERIFY_FAIL++
        Write-HostState -Path $statePath -State $state
    }
}

# --- Итоги ---
Write-Log ""
Write-Log "=========================================="
Write-Log "=== ИТОГО ==="
Write-Log "=========================================="
Write-Log "Всего хостов:        $($filteredHosts.Count)"
Write-Log "OK (всего):          $($stats.OK)"
Write-Log "  из них уже было:   $($stats.ALREADY_WORKING)"
Write-Log "Skipped (state OK):  $($stats.SKIPPED)"
Write-Log "Offline:             $($stats.OFFLINE)"
Write-Log "SMB unreachable:     $($stats.SMB_FAIL)"
Write-Log "PsExec failed:       $($stats.PSEXEC_FAIL)"
Write-Log "Verify failed:       $($stats.VERIFY_FAIL)"
Write-Log ""
Write-Log "State: $stateDir"
Write-Log "Лог: $LogFile"

$failures = $stats.PSEXEC_FAIL + $stats.VERIFY_FAIL
if ($failures -gt 0) { exit 1 }

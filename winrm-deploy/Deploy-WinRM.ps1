<#
.SYNOPSIS
    Массовая настройка WinRM на Windows-хостах через PsExec.
.DESCRIPTION
    Скрипт читает список Windows-хостов из inventory.yaml,
    проверяет доступность WinRM, и если недоступен — настраивает
    WinRM удалённо через PsExec.exe.
.PARAMETER DryRun
    Только показать список хостов и статусы, без изменений.
.PARAMETER HostFilter
    Фильтр по имени хоста или IP (поддерживает wildcard *).
.PARAMETER ConfigPath
    Путь к config.yaml. По умолчанию — рядом со скриптом.
.EXAMPLE
    .\Deploy-WinRM.ps1 -DryRun
    .\Deploy-WinRM.ps1 -HostFilter "buh01-ws"
    .\Deploy-WinRM.ps1
#>

param(
    [switch]$DryRun,
    [string]$HostFilter = "*",
    [string]$ConfigPath
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$LogFile = Join-Path $ScriptDir "Deploy-WinRM.log"

# ============================================================
# Вспомогательные функции
# ============================================================

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        "ERROR" { Write-Host $entry -ForegroundColor Red }
        "WARN"  { Write-Host $entry -ForegroundColor Yellow }
        "OK"    { Write-Host $entry -ForegroundColor Green }
        default { Write-Host $entry }
    }
    Add-Content -Path $LogFile -Value $entry -Encoding UTF8 -ErrorAction SilentlyContinue
}

function Parse-SimpleYaml {
    <#
    .SYNOPSIS
        Минималистичный парсер YAML → вложенный hashtable.
        Обрабатывает только mapping (key: value) и вложенность по отступам.
    #>
    param([string]$Path)

    $lines = Get-Content $Path
    $root = @{}
    $stack = @( @{ obj = $root; indent = -1 } )

    foreach ($line in $lines) {
        # Пропускаем пустые строки и комментарии
        if ($line -match '^\s*$' -or $line -match '^\s*#') { continue }

        # Считаем отступ
        $trimmed = $line.TrimStart()
        $indent = $line.Length - $trimmed.Length

        # Убираем уровни стека до текущего отступа
        while ($stack.Count -gt 1 -and $stack[-1].indent -ge $indent) {
            $stack = $stack[0..($stack.Count - 2)]
        }

        $current = $stack[-1].obj

        if ($trimmed -match '^(.+?):\s*(.+)$') {
            # key: value
            $key = $Matches[1].Trim()
            $val = $Matches[2].Trim().Trim('"').Trim("'")

            # Обработка {} как пустого объекта
            if ($val -eq '{}') {
                $current[$key] = @{}
            } else {
                $current[$key] = $val
            }
        }
        elseif ($trimmed -match '^(.+?):\s*$') {
            # key: (вложенный объект)
            $key = $Matches[1].Trim()
            $newObj = @{}
            $current[$key] = $newObj
            $stack += @{ obj = $newObj; indent = $indent }
        }
    }

    return $root
}

function Get-WindowsHosts {
    <#
    .SYNOPSIS
        Извлекает список Windows-хостов из inventory.yaml.
        Возвращает массив @{ Name; IP }.
    #>
    param([hashtable]$Inventory)

    $hosts = @()
    $windowsGroup = $Inventory['all']['children']['windows']
    if (-not $windowsGroup) { return $hosts }

    # Обходим children (windows_servers, windows_workstations)
    if ($windowsGroup['children']) {
        foreach ($subGroupName in $windowsGroup['children'].Keys) {
            $subGroup = $windowsGroup['children'][$subGroupName]
            if ($subGroup['hosts']) {
                foreach ($hostName in $subGroup['hosts'].Keys) {
                    $hostData = $subGroup['hosts'][$hostName]
                    if ($hostData -is [hashtable] -and $hostData['ansible_host']) {
                        $hosts += @{
                            Name = $hostName
                            IP   = $hostData['ansible_host']
                        }
                    }
                }
            }
        }
    }

    # Прямые хосты в windows.hosts
    if ($windowsGroup['hosts'] -and $windowsGroup['hosts'] -is [hashtable]) {
        foreach ($hostName in $windowsGroup['hosts'].Keys) {
            $hostData = $windowsGroup['hosts'][$hostName]
            if ($hostData -is [hashtable] -and $hostData['ansible_host']) {
                $hosts += @{
                    Name = $hostName
                    IP   = $hostData['ansible_host']
                }
            }
        }
    }

    return $hosts
}

function Test-WinRM {
    param([string]$ComputerName, [int]$TimeoutSec = 5)

    try {
        # Сначала быстрая проверка порта
        $tcp = New-Object System.Net.Sockets.TcpClient
        $connect = $tcp.BeginConnect($ComputerName, 5985, $null, $null)
        $wait = $connect.AsyncWaitHandle.WaitOne($TimeoutSec * 1000, $false)
        if (-not $wait) {
            $tcp.Close()
            return $false
        }
        $tcp.EndConnect($connect)
        $tcp.Close()

        # Порт открыт — проверяем WinRM
        $result = Test-WSMan -ComputerName $ComputerName -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Invoke-PsExecWinRM {
    <#
    .SYNOPSIS
        Настраивает WinRM на удалённом хосте через PsExec.
        Выполняет один вызов PowerShell — избегает проблем с экранированием кавычек.
    #>
    param(
        [string]$PsExecPath,
        [string]$ComputerName,
        [string]$Username,
        [string]$Password
    )

    # Один PowerShell-скрипт, который выполнится на удалённой машине
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

    Write-Log "  PsExec -> \\$ComputerName : Enable-PSRemoting + WinRM config"

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

    $stdout = if (Test-Path "$env:TEMP\psexec_out.txt") {
        Get-Content "$env:TEMP\psexec_out.txt" -Raw
    } else { "" }
    $stderr = if (Test-Path "$env:TEMP\psexec_err.txt") {
        Get-Content "$env:TEMP\psexec_err.txt" -Raw
    } else { "" }

    $result = @{
        ExitCode = $proc.ExitCode
        StdOut   = $stdout
        StdErr   = $stderr
    }

    if ($proc.ExitCode -eq 0) {
        Write-Log "  Настройка завершена успешно" "OK"
    } else {
        Write-Log "  Ошибка (exitcode=$($proc.ExitCode)): $stderr" "ERROR"
    }

    return $result
}

# ============================================================
# MAIN
# ============================================================

Write-Log "=========================================="
Write-Log "=== Deploy-WinRM — Начало ==="
Write-Log "=========================================="

# --- 1. Загрузка конфига ---
if (-not $ConfigPath) {
    $ConfigPath = Join-Path $ScriptDir "config.yaml"
}

if (-not (Test-Path $ConfigPath)) {
    Write-Log "Конфиг не найден: $ConfigPath" "ERROR"
    exit 1
}

Write-Log "Конфиг: $ConfigPath"
$config = Parse-SimpleYaml -Path $ConfigPath

$psexecPath = $config['psexec_path']
$inventoryPath = $config['inventory_path']
$username = $config['credentials']['username']
$winrmTimeout = [int]($config['winrm_test_timeout'])

if (-not (Test-Path $psexecPath)) {
    Write-Log "PsExec не найден: $psexecPath" "ERROR"
    exit 1
}

if (-not (Test-Path $inventoryPath)) {
    Write-Log "Inventory не найден: $inventoryPath" "ERROR"
    exit 1
}

# --- 2. Получение пароля ---
if (-not $DryRun) {
    $cfgPassword = $config['credentials']['password']
    if ($cfgPassword -and $cfgPassword -ne 'PUT_PASSWORD_HERE') {
        Write-Log "Пароль: из конфига"
        $password = $cfgPassword
    } else {
        Write-Log "Запрос учётных данных для $username..."
        $cred = Get-Credential -UserName $username -Message "Пароль для PsExec ($username)"
        if (-not $cred) {
            Write-Log "Отменено пользователем" "ERROR"
            exit 1
        }
        $password = $cred.GetNetworkCredential().Password
    }
}

# --- 3. Загрузка inventory ---
Write-Log "Чтение inventory: $inventoryPath"
$inventory = Parse-SimpleYaml -Path $inventoryPath
$allHosts = Get-WindowsHosts -Inventory $inventory

# Фильтрация
$filteredHosts = $allHosts | Where-Object {
    $_.Name -like $HostFilter -or $_.IP -like $HostFilter
}

# Сортировка по IP
$filteredHosts = $filteredHosts | Sort-Object { [version]($_.IP -replace '(\d+)\.(\d+)\.(\d+)\.(\d+)', '$1.$2.$3.$4') }

Write-Log "Всего Windows-хостов: $($allHosts.Count)"
Write-Log "После фильтра '$HostFilter': $($filteredHosts.Count)"

if ($filteredHosts.Count -eq 0) {
    Write-Log "Нет хостов для обработки" "WARN"
    exit 0
}

# --- 4. Фаза 1: Проверка WinRM ---
Write-Log ""
Write-Log "=== Фаза 1: Проверка WinRM ==="
Write-Log ""

$winrmOK = @()
$winrmFail = @()
$unreachable = @()

foreach ($h in $filteredHosts) {
    $displayName = "$($h.Name) ($($h.IP))"

    # Быстрая проверка — хост вообще доступен? (ping)
    $ping = Test-Connection -ComputerName $h.IP -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $ping) {
        Write-Log "  [ OFFLINE ] $displayName" "WARN"
        $unreachable += $h
        continue
    }

    $winrmStatus = Test-WinRM -ComputerName $h.IP -TimeoutSec $winrmTimeout
    if ($winrmStatus) {
        Write-Log "  [  OK  ] $displayName — WinRM работает" "OK"
        $winrmOK += $h
    } else {
        Write-Log "  [ FAIL ] $displayName — WinRM недоступен"
        $winrmFail += $h
    }
}

Write-Log ""
Write-Log "Итого фазы 1: OK=$($winrmOK.Count)  FAIL=$($winrmFail.Count)  OFFLINE=$($unreachable.Count)"

if ($DryRun) {
    Write-Log ""
    Write-Log "=== DryRun — без изменений ==="
    Write-Log "Хосты для настройки WinRM:"
    foreach ($h in $winrmFail) {
        Write-Log "  - $($h.Name) ($($h.IP))"
    }
    exit 0
}

if ($winrmFail.Count -eq 0) {
    Write-Log "Все хосты уже с WinRM — настройка не нужна" "OK"
    exit 0
}

# --- 5. Фаза 2: Настройка WinRM через PsExec ---
Write-Log ""
Write-Log "=== Фаза 2: Настройка WinRM через PsExec ==="
Write-Log ""

$configuredOK = @()
$configuredFail = @()

foreach ($h in $winrmFail) {
    $displayName = "$($h.Name) ($($h.IP))"
    Write-Log "Настройка: $displayName"

    try {
        $result = Invoke-PsExecWinRM `
            -PsExecPath $psexecPath `
            -ComputerName $h.IP `
            -Username $username `
            -Password $password

        if ($result.ExitCode -eq 0) {
            $configuredOK += $h
        } else {
            $configuredFail += $h
        }
    }
    catch {
        Write-Log "  Исключение для $displayName : $_" "ERROR"
        $configuredFail += $h
    }
}

# --- 6. Фаза 3: Повторная проверка WinRM ---
Write-Log ""
Write-Log "=== Фаза 3: Повторная проверка WinRM ==="
Write-Log ""

# Даём время сервисам подняться
Start-Sleep -Seconds 3

$verifyOK = @()
$verifyFail = @()

foreach ($h in $configuredOK) {
    $displayName = "$($h.Name) ($($h.IP))"
    $winrmStatus = Test-WinRM -ComputerName $h.IP -TimeoutSec $winrmTimeout

    if ($winrmStatus) {
        Write-Log "  [  OK  ] $displayName — WinRM работает" "OK"
        $verifyOK += $h
    } else {
        Write-Log "  [ FAIL ] $displayName — WinRM всё ещё недоступен" "ERROR"
        $verifyFail += $h
    }
}

# --- 7. Итоговый отчёт ---
Write-Log ""
Write-Log "=========================================="
Write-Log "=== ИТОГО ==="
Write-Log "=========================================="
Write-Log "Всего хостов:          $($filteredHosts.Count)"
Write-Log "WinRM уже работал:     $($winrmOK.Count)"
Write-Log "Офлайн:                $($unreachable.Count)"
Write-Log "Настроено успешно:     $($verifyOK.Count)"
Write-Log "Настроено с ошибками:  $($configuredFail.Count)"
Write-Log "Не прошло проверку:    $($verifyFail.Count)"

if ($unreachable.Count -gt 0) {
    Write-Log ""
    Write-Log "Офлайн хосты:"
    foreach ($h in $unreachable) {
        Write-Log "  - $($h.Name) ($($h.IP))" "WARN"
    }
}

if ($configuredFail.Count -gt 0 -or $verifyFail.Count -gt 0) {
    Write-Log ""
    Write-Log "Проблемные хосты:"
    foreach ($h in ($configuredFail + $verifyFail)) {
        Write-Log "  - $($h.Name) ($($h.IP))" "ERROR"
    }
    exit 1
}

Write-Log ""
Write-Log "=== Готово ===" "OK"
Write-Log "Лог: $LogFile"

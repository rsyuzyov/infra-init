<#
.SYNOPSIS
    Массовая установка и настройка OpenSSH Server на Windows-хостах.
.DESCRIPTION
    Транспорты по приоритету (для каждого хоста):
      1. WinRM (Invoke-Command) — основной, быстрый, без admin$
      2. PsExec — fallback, если WinRM недоступен
      3. MSI fallback — внутри обоих транспортов, для старых Windows без Capability API

    Per-host state: output/<domain>/openssh/<host>.json.
    Уже OK-хосты пропускаются; -Force переобрабатывает.

.PARAMETER DryRun
    Только показать план, без изменений.
.PARAMETER HostFilter
    Wildcard-фильтр по имени или IP.
.PARAMETER HostsFile
    Файл со списком хостов, по одному на строку (приоритетнее HostFilter).
.PARAMETER ConfigPath
    Путь к config.yaml. По умолчанию — рядом со скриптом.
.PARAMETER Force
    Игнорировать сохранённый state.

.EXAMPLE
    .\Deploy-OpenSSH.ps1 -DryRun
    .\Deploy-OpenSSH.ps1 -HostsFile .\tmp\retry-hosts-openssh.txt
    .\Deploy-OpenSSH.ps1 -Force
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
. (Join-Path $ScriptDir "Common.ps1")

Assert-Admin -ScriptPath $MyInvocation.MyCommand.Definition -BoundParams $PSBoundParameters `
    -SkipIf { Test-WSManClientReady } -Reason "настройка локального WSMan-клиента"

$LogFile = New-LogFile -RootDir $ScriptDir -ScriptName "Deploy-OpenSSH"

# ============================================================
# Установочный scriptblock для WinRM (нативный, возвращает hashtable)
# ============================================================

$OpenSshInstallScript = {
    param($sshPort, $pwdAuth, $pubkeyAuth, $defaultShell)
    $ErrorActionPreference = 'Stop'
    $result = @{
        status = 'ERROR'; category = ''; message = ''; sshd_status = ''; restart_needed = $false
    }
    try {
        if (-not (Get-Command -Name Get-WindowsCapability -ErrorAction SilentlyContinue)) {
            $result.status = 'NO_CAPABILITY_API'
            $result.category = 'NO_CAPABILITY_API'
            $result.message = 'Get-WindowsCapability not available — use MSI fallback'
            return $result
        }
        $feat = Get-WindowsCapability -Online | Where-Object { $_.Name -like 'OpenSSH.Server*' }
        if (-not $feat) {
            $result.status = 'NO_CAPABILITY_API'
            $result.category = 'NO_CAPABILITY_API'
            $result.message = 'OpenSSH.Server capability not listed — use MSI fallback'
            return $result
        }

        if ($feat.State -ne 'Installed') {
            $r = Add-WindowsCapability -Online -Name 'OpenSSH.Server~~~~0.0.1.0' -ErrorAction Stop
            if ($r.RestartNeeded) {
                $result.status = 'REBOOT_REQUIRED'
                $result.category = 'REBOOT_REQUIRED'
                $result.message = 'Add-WindowsCapability requires reboot'
                $result.restart_needed = $true
                return $result
            }
        }

        $cfgDir = $env:ProgramData + '\ssh'
        $cfgPath = $cfgDir + '\sshd_config'
        if (-not (Test-Path $cfgDir)) {
            $result.status = 'REBOOT_REQUIRED'
            $result.category = 'REBOOT_REQUIRED'
            $result.message = "Capability accepted but $cfgDir not created — pending reboot"
            $result.restart_needed = $true
            return $result
        }

        if (Test-Path $cfgPath) {
            $bak = $cfgPath + '.bak.' + (Get-Date -Format 'yyyyMMdd_HHmmss')
            Copy-Item $cfgPath $bak -ErrorAction Stop
        }

        $cfg = @"
Port $sshPort
AddressFamily any
ListenAddress 0.0.0.0
PubkeyAuthentication $pubkeyAuth
PasswordAuthentication $pwdAuth
PermitEmptyPasswords no
MaxAuthTries 5
MaxSessions 10
LoginGraceTime 60
SyslogFacility LOCAL0
LogLevel INFO
Subsystem sftp sftp-server.exe
Match Group administrators
    AuthorizedKeysFile __PROGRAMDATA__/ssh/administrators_authorized_keys
"@
        Set-Content -Path $cfgPath -Value $cfg -Encoding UTF8 -ErrorAction Stop

        $regPath = 'HKLM:\SOFTWARE\OpenSSH'
        if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force -ErrorAction Stop | Out-Null }
        Set-ItemProperty -Path $regPath -Name 'DefaultShell' -Value $defaultShell -ErrorAction Stop
        Set-ItemProperty -Path $regPath -Name 'DefaultShellCommandOption' -Value '/c' -ErrorAction Stop

        $fw = Get-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -ErrorAction SilentlyContinue
        if (-not $fw) {
            New-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -DisplayName 'OpenSSH Server (SSH)' `
                -Enabled True -Direction Inbound -Protocol TCP -Action Allow -LocalPort $sshPort `
                -ErrorAction Stop | Out-Null
        } else {
            Enable-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -ErrorAction SilentlyContinue
        }

        $svc = Get-Service -Name sshd -ErrorAction Stop
        Set-Service -Name sshd -StartupType Automatic -ErrorAction Stop
        Restart-Service sshd -Force -ErrorAction Stop
        Start-Sleep -Seconds 2
        $svc = Get-Service -Name sshd -ErrorAction Stop
        $result.sshd_status = $svc.Status.ToString()

        if ($svc.Status -eq 'Running') {
            $result.status = 'OK'
            $result.category = 'INSTALLED'
            $result.message = 'sshd configured and running'
        } else {
            $result.status = 'FAIL'
            $result.category = 'SSHD_NOT_RUNNING'
            $result.message = "sshd status: $($svc.Status)"
        }
    } catch {
        $result.status = 'FAIL'
        $result.category = 'EXCEPTION'
        $result.message = $_.Exception.Message
    }
    return $result
}

# MSI-installer scriptblock (для WinRM-варианта). Требует чтобы MSI уже был на хосте.
$OpenSshInstallMsiScript = {
    param($msiLocalPath, $sshPort, $pwdAuth, $pubkeyAuth, $defaultShell)
    $ErrorActionPreference = 'Stop'
    $result = @{ status = 'ERROR'; category = ''; message = ''; sshd_status = '' }
    try {
        $msiArgs = @('/i', $msiLocalPath, '/qn', '/norestart')
        $p = Start-Process -FilePath 'msiexec.exe' -ArgumentList $msiArgs -Wait -PassThru
        if ($p.ExitCode -ne 0 -and $p.ExitCode -ne 3010) {
            $result.status = 'FAIL'; $result.category = 'MSI_INSTALL_FAILED'
            $result.message = "msiexec exitcode $($p.ExitCode)"
            return $result
        }

        $svc = Get-Service -Name sshd -ErrorAction SilentlyContinue
        if (-not $svc) {
            $installScript = 'C:\Program Files\OpenSSH\install-sshd.ps1'
            if (Test-Path $installScript) {
                & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $installScript | Out-Null
            } else {
                $result.status = 'FAIL'; $result.category = 'INSTALL_SCRIPT_MISSING'
                $result.message = 'install-sshd.ps1 not found'
                return $result
            }
            $svc = Get-Service -Name sshd -ErrorAction Stop
        }
        Set-Service -Name sshd -StartupType Automatic -ErrorAction Stop
        if ($svc.Status -ne 'Running') {
            Start-Service sshd -ErrorAction Stop
            Start-Sleep -Seconds 2
        }

        $cfgDir = $env:ProgramData + '\ssh'
        $cfgPath = $cfgDir + '\sshd_config'
        if (-not (Test-Path $cfgDir)) {
            New-Item -ItemType Directory -Path $cfgDir -Force -ErrorAction Stop | Out-Null
        }
        if (Test-Path $cfgPath) {
            $bak = $cfgPath + '.bak.' + (Get-Date -Format 'yyyyMMdd_HHmmss')
            Copy-Item $cfgPath $bak -ErrorAction Stop
        }
        $cfg = @"
Port $sshPort
AddressFamily any
ListenAddress 0.0.0.0
PubkeyAuthentication $pubkeyAuth
PasswordAuthentication $pwdAuth
PermitEmptyPasswords no
MaxAuthTries 5
MaxSessions 10
LoginGraceTime 60
SyslogFacility LOCAL0
LogLevel INFO
Subsystem sftp sftp-server.exe
Match Group administrators
    AuthorizedKeysFile __PROGRAMDATA__/ssh/administrators_authorized_keys
"@
        Set-Content -Path $cfgPath -Value $cfg -Encoding UTF8 -ErrorAction Stop

        $regPath = 'HKLM:\SOFTWARE\OpenSSH'
        if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force -ErrorAction Stop | Out-Null }
        Set-ItemProperty -Path $regPath -Name 'DefaultShell' -Value $defaultShell -ErrorAction Stop
        Set-ItemProperty -Path $regPath -Name 'DefaultShellCommandOption' -Value '/c' -ErrorAction Stop

        $fw = Get-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -ErrorAction SilentlyContinue
        if (-not $fw) {
            if (Get-Command -Name New-NetFirewallRule -ErrorAction SilentlyContinue) {
                New-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -DisplayName 'OpenSSH Server (SSH)' `
                    -Enabled True -Direction Inbound -Protocol TCP -Action Allow -LocalPort $sshPort -ErrorAction Stop | Out-Null
            } else {
                & netsh.exe advfirewall firewall add rule name="OpenSSH-Server-In-TCP" dir=in action=allow protocol=TCP localport=$sshPort | Out-Null
            }
        } else {
            Enable-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -ErrorAction SilentlyContinue
        }

        Restart-Service sshd -Force -ErrorAction Stop
        Start-Sleep -Seconds 2
        $svc = Get-Service -Name sshd -ErrorAction Stop
        $result.sshd_status = $svc.Status.ToString()

        if ($svc.Status -eq 'Running') {
            $result.status = 'OK'; $result.category = 'INSTALLED_MSI'
            $result.message = 'sshd configured and running (MSI)'
        } else {
            $result.status = 'FAIL'; $result.category = 'SSHD_NOT_RUNNING'
            $result.message = "sshd status after MSI: $($svc.Status)"
        }
    } catch {
        $result.status = 'FAIL'; $result.category = 'EXCEPTION'
        $result.message = $_.Exception.Message
    } finally {
        Remove-Item -Path $msiLocalPath -Force -ErrorAction SilentlyContinue
    }
    return $result
}

# ============================================================
# Транспорт: WinRM (Invoke-Command)
# ============================================================

function Invoke-WinRmInstall {
    param(
        [string]$ComputerName,
        [PSCredential]$Credential,
        [hashtable]$SshdSettings,
        [string]$AuthMethod = 'Default'
    )
    try {
        $session = New-PSSession -ComputerName $ComputerName -Credential $Credential `
            -Authentication $AuthMethod -ErrorAction Stop
    } catch {
        return @{ status = 'WINRM_CONNECT_FAIL'; category = 'WINRM_CONNECT_FAIL'; message = $_.Exception.Message }
    }
    try {
        $res = Invoke-Command -Session $session -ScriptBlock $OpenSshInstallScript `
            -ArgumentList @(
                $SshdSettings['port'],
                $SshdSettings['password_authentication'],
                $SshdSettings['pubkey_authentication'],
                $SshdSettings['default_shell']
            ) -ErrorAction Stop
        return $res
    } catch {
        return @{ status = 'WINRM_INVOKE_FAIL'; category = 'WINRM_INVOKE_FAIL'; message = $_.Exception.Message }
    } finally {
        Remove-PSSession $session -ErrorAction SilentlyContinue
    }
}

function Invoke-WinRmMsiInstall {
    param(
        [string]$ComputerName,
        [PSCredential]$Credential,
        [string]$MsiLocalPath,    # путь к MSI на локальной машине
        [hashtable]$SshdSettings,
        [string]$AuthMethod = 'Default'
    )
    $session = $null
    try {
        $session = New-PSSession -ComputerName $ComputerName -Credential $Credential `
            -Authentication $AuthMethod -ErrorAction Stop
        # Копируем MSI на хост через PSSession (без admin$/SMB)
        $remoteMsi = "C:\Windows\Temp\$([System.IO.Path]::GetFileName($MsiLocalPath))"
        Copy-Item -Path $MsiLocalPath -Destination $remoteMsi -ToSession $session -Force -ErrorAction Stop

        $res = Invoke-Command -Session $session -ScriptBlock $OpenSshInstallMsiScript `
            -ArgumentList @(
                $remoteMsi,
                $SshdSettings['port'],
                $SshdSettings['password_authentication'],
                $SshdSettings['pubkey_authentication'],
                $SshdSettings['default_shell']
            ) -ErrorAction Stop
        return $res
    } catch {
        return @{ status = 'WINRM_MSI_FAIL'; category = 'WINRM_MSI_FAIL'; message = $_.Exception.Message }
    } finally {
        if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
    }
}

# ============================================================
# Транспорт: PsExec (fallback, текущая логика)
# ============================================================

function Invoke-PsExecOpenSSH {
    param(
        [string]$PsExecPath,
        [string]$ComputerName,
        [string]$Username,
        [string]$Password,
        [hashtable]$SshdSettings
    )
    $sshPort      = $SshdSettings['port']
    $pwdAuth      = if ($SshdSettings['password_authentication'] -eq 'yes') { 'yes' } else { 'no' }
    $pubkeyAuth   = if ($SshdSettings['pubkey_authentication']   -eq 'yes') { 'yes' } else { 'no' }
    $defaultShell = $SshdSettings['default_shell']

    $remoteScript = @"
`$ErrorActionPreference = 'Stop'
try {
    if (-not (Get-Command -Name Get-WindowsCapability -ErrorAction SilentlyContinue)) {
        Write-Host 'NO_CAPABILITY_API'; exit 2
    }
    `$feat = Get-WindowsCapability -Online | Where-Object { `$_.Name -like 'OpenSSH.Server*' }
    if (-not `$feat) { Write-Host 'NO_CAPABILITY_API'; exit 2 }
    if (`$feat.State -ne 'Installed') {
        `$r = Add-WindowsCapability -Online -Name 'OpenSSH.Server~~~~0.0.1.0' -ErrorAction Stop
        if (`$r.RestartNeeded) { Write-Host 'REBOOT_REQUIRED'; exit 3 }
    }
    `$cfgDir = `$env:ProgramData + '\ssh'
    `$cfgPath = `$cfgDir + '\sshd_config'
    if (-not (Test-Path `$cfgDir)) { Write-Host 'REBOOT_REQUIRED'; exit 3 }
    if (Test-Path `$cfgPath) {
        `$bak = `$cfgPath + '.bak.' + (Get-Date -Format 'yyyyMMdd_HHmmss')
        Copy-Item `$cfgPath `$bak -ErrorAction Stop
    }
    `$cfg = @'
Port $sshPort
AddressFamily any
ListenAddress 0.0.0.0
PubkeyAuthentication $pubkeyAuth
PasswordAuthentication $pwdAuth
PermitEmptyPasswords no
MaxAuthTries 5
MaxSessions 10
LoginGraceTime 60
SyslogFacility LOCAL0
LogLevel INFO
Subsystem sftp sftp-server.exe
Match Group administrators
    AuthorizedKeysFile __PROGRAMDATA__/ssh/administrators_authorized_keys
'@
    Set-Content -Path `$cfgPath -Value `$cfg -Encoding UTF8 -ErrorAction Stop
    `$regPath = 'HKLM:\SOFTWARE\OpenSSH'
    if (-not (Test-Path `$regPath)) { New-Item -Path `$regPath -Force -ErrorAction Stop | Out-Null }
    Set-ItemProperty -Path `$regPath -Name 'DefaultShell' -Value '$defaultShell' -ErrorAction Stop
    Set-ItemProperty -Path `$regPath -Name 'DefaultShellCommandOption' -Value '/c' -ErrorAction Stop
    `$fw = Get-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -ErrorAction SilentlyContinue
    if (-not `$fw) {
        New-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -DisplayName 'OpenSSH Server (SSH)' -Enabled True -Direction Inbound -Protocol TCP -Action Allow -LocalPort $sshPort -ErrorAction Stop | Out-Null
    } else {
        Enable-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -ErrorAction SilentlyContinue
    }
    `$svc = Get-Service -Name sshd -ErrorAction Stop
    Set-Service -Name sshd -StartupType Automatic -ErrorAction Stop
    Restart-Service sshd -Force -ErrorAction Stop
    Start-Sleep -Seconds 2
    `$svc = Get-Service -Name sshd -ErrorAction Stop
    if (`$svc.Status -eq 'Running') { Write-Host 'OPENSSH_OK' } else { Write-Host 'OPENSSH_FAIL'; exit 1 }
} catch {
    Write-Host "OPENSSH_ERROR: `$(`$_.Exception.Message)"; exit 1
}
"@

    $psexecArgs = @(
        "\\$ComputerName", "-u", $Username, "-p", $Password, "-accepteula", "-nobanner", "-h",
        "powershell.exe", "-ExecutionPolicy", "Bypass", "-Command", $remoteScript
    )
    $proc = Start-Process -FilePath $PsExecPath -ArgumentList $psexecArgs `
        -NoNewWindow -Wait -PassThru `
        -RedirectStandardOutput "$env:TEMP\psexec_ssh_out.txt" `
        -RedirectStandardError  "$env:TEMP\psexec_ssh_err.txt"
    $stdout = if (Test-Path "$env:TEMP\psexec_ssh_out.txt") { Get-Content "$env:TEMP\psexec_ssh_out.txt" -Raw } else { '' }
    $stderr = if (Test-Path "$env:TEMP\psexec_ssh_err.txt") { Get-Content "$env:TEMP\psexec_ssh_err.txt" -Raw } else { '' }

    $result = @{ status = 'FAIL'; category = ''; message = ''; exit_code = $proc.ExitCode }
    if ($proc.ExitCode -eq 0)      { $result.status = 'OK';              $result.category = 'INSTALLED_VIA_PSEXEC';   $result.message = 'PsExec configured sshd' }
    elseif ($proc.ExitCode -eq 2)  { $result.status = 'NO_CAPABILITY_API'; $result.category = 'NO_CAPABILITY_API';     $result.message = 'No Capability API — try MSI' }
    elseif ($proc.ExitCode -eq 3)  { $result.status = 'REBOOT_REQUIRED'; $result.category = 'REBOOT_REQUIRED';         $result.message = 'Pending reboot' }
    elseif ($proc.ExitCode -eq 6)  { $result.status = 'SMB_FAIL';        $result.category = 'SMB_UNREACHABLE';         $result.message = 'PsExec exitcode 6 (admin$ unreachable)' }
    else                            { $result.status = 'PSEXEC_FAIL';     $result.category = 'PSEXEC_ERROR';            $result.message = $stderr.Trim() }
    return $result
}

# ============================================================
# Транспорт-выбор: WinRM сначала, PsExec потом
# ============================================================

function Install-OpenSSH {
    # Возвращает hashtable с полями: status, category, message, transport, exit_code, sshd_status
    param(
        $Target,               # @{ Name=; IP= }  — не используем имя $Host: оно зарезервировано в PS
        [PSCredential]$Credential,
        [string]$Username,
        [string]$Password,
        [hashtable]$SshdSettings,
        [string]$PsExecPath,
        [string]$MsiPath,
        [bool]$MsiAvailable,
        [int]$WinrmTimeout = 5,
        [int]$SmbTimeout = 5
    )
    $ip = $Target.IP

    # 1) Пробуем WinRM, если порт 5985 открыт.
    if (Test-TcpPort -ComputerName $ip -Port 5985 -TimeoutSec $WinrmTimeout) {
        Write-Log "  [winrm] $($Target.Name) — пробуем установку через WinRM"
        $r = Invoke-WinRmInstall -ComputerName $ip -Credential $Credential -SshdSettings $SshdSettings

        if ($r.status -eq 'OK' -or $r.status -eq 'REBOOT_REQUIRED') {
            $r['transport'] = 'winrm'
            return $r
        }
        if ($r.status -eq 'NO_CAPABILITY_API' -and $MsiAvailable) {
            Write-Log "  [winrm] нет Capability API — fallback на MSI через WinRM"
            $r2 = Invoke-WinRmMsiInstall -ComputerName $ip -Credential $Credential `
                -MsiLocalPath $MsiPath -SshdSettings $SshdSettings
            $r2['transport'] = 'winrm-msi'
            return $r2
        }
        # WinRM не доехал по сетевой причине → пробуем PsExec
        if ($r.status -eq 'WINRM_CONNECT_FAIL' -or $r.status -eq 'WINRM_INVOKE_FAIL') {
            Write-Log "  [winrm] упал ($($r.category)): $($r.message). Fallback на PsExec." "WARN"
        } else {
            # WinRM сработал, но не успешно (FAIL, NO_CAPABILITY_API без MSI) — отдаём это как финальный
            $r['transport'] = 'winrm'
            return $r
        }
    } else {
        Write-Log "  [winrm] $($Target.Name) — TCP 5985 закрыт, переходим к PsExec" "WARN"
    }

    # 2) PsExec — требует TCP 445
    if (-not (Test-TcpPort -ComputerName $ip -Port 445 -TimeoutSec $SmbTimeout)) {
        return @{
            status = 'SMB_FAIL'; category = 'SMB_UNREACHABLE'
            message = 'WinRM unavailable, TCP 445 also closed'
            transport = 'none'
        }
    }
    $r = Invoke-PsExecOpenSSH -PsExecPath $PsExecPath -ComputerName $ip `
        -Username $Username -Password $Password -SshdSettings $SshdSettings
    $r['transport'] = 'psexec'

    if ($r.status -eq 'NO_CAPABILITY_API' -and $MsiAvailable) {
        # MSI через PsExec — не реализуем; в логе только подсказываем
        $r['message'] += '. WinRM-MSI tried earlier failed; MSI via PsExec not implemented in this path.'
    }
    return $r
}

# ============================================================
# MAIN
# ============================================================

Write-Log "=========================================="
Write-Log "=== Deploy-OpenSSH — Начало ==="
Write-Log "=========================================="
Write-Log "Лог: $LogFile"

if (-not $ConfigPath) { $ConfigPath = Join-Path $ScriptDir "config.yaml" }
if (-not (Test-Path $ConfigPath)) {
    Write-Log "Конфиг не найден: $ConfigPath" "ERROR"; exit 1
}
Write-Log "Конфиг: $ConfigPath"
$config = Parse-SimpleYaml -Path $ConfigPath

$psexecPath    = $config['psexec_path']
$inventoryPath = $config['inventory_path']
$msiPath       = $config['openssh_msi_path']
$username      = $config['credentials']['username']
$sshTestTimeout = if ($config['ssh_test_timeout']) { [int]$config['ssh_test_timeout'] } else { 5 }
$domain        = if ($config['domain']) { $config['domain'] } else { '_default' }

$sshdSettings = @{
    port                     = if ($config['sshd'] -and $config['sshd']['port']) { $config['sshd']['port'] } else { '22' }
    password_authentication  = if ($config['sshd'] -and $config['sshd']['password_authentication']) { $config['sshd']['password_authentication'] } else { 'yes' }
    pubkey_authentication    = if ($config['sshd'] -and $config['sshd']['pubkey_authentication']) { $config['sshd']['pubkey_authentication'] } else { 'yes' }
    default_shell            = if ($config['sshd'] -and $config['sshd']['default_shell']) { $config['sshd']['default_shell'] } else { 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' }
}
$sshPort = [int]$sshdSettings['port']

if (-not (Test-Path $psexecPath)) { Write-Log "PsExec не найден: $psexecPath" "ERROR"; exit 1 }
if (-not (Test-Path $inventoryPath)) { Write-Log "Inventory не найден: $inventoryPath" "ERROR"; exit 1 }

$msiAvailable = $false
if ($msiPath -and (Test-Path $msiPath)) {
    $msiAvailable = $true
    Write-Log "MSI fallback доступен: $msiPath"
} else {
    Write-Log "MSI fallback недоступен (нет файла или путь не задан)" "WARN"
}

Write-Log "Domain: $domain"

if (Test-WSManClientReady) {
    Write-Log "Локальный WSMan: TrustedHosts=*, Basic+AllowUnencrypted включены"
} elseif (Test-IsAdmin -and (Initialize-LocalWSMan)) {
    Write-Log "Локальный WSMan донастроен (admin)"
} else {
    Write-Log "Локальный WSMan не настроен — WinRM-транспорт работать не будет. Запусти .\Setup-LocalEnvironment.ps1 от админа." "WARN"
}

# --- Пароль ---
$password = $null
$cred = $null
if (-not $DryRun) {
    $cfgPassword = $config['credentials']['password']
    if ($cfgPassword -and $cfgPassword -ne 'PUT_PASSWORD_HERE') {
        $password = $cfgPassword
        Write-Log "Пароль: из конфига"
    } else {
        $c = Get-Credential -UserName $username -Message "Пароль для $username"
        if (-not $c) { Write-Log "Отменено пользователем" "ERROR"; exit 1 }
        $password = $c.GetNetworkCredential().Password
    }
    $secPwd = ConvertTo-SecureString $password -AsPlainText -Force
    $cred = New-Object System.Management.Automation.PSCredential($username, $secPwd)
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
if ($filteredHosts.Count -eq 0) { Write-Log "Нет хостов для обработки" "WARN"; exit 0 }

$stateDir = Get-StateDir -RootDir $ScriptDir -Domain $domain -Transport 'openssh'
Write-Log "State directory: $stateDir"
if ($Force) { Write-Log "Force: state будет игнорироваться" "WARN" }

Write-Log ""
Write-Log "=== Обработка ==="
Write-Log ""

$stats = @{
    OK = 0; ALREADY_WORKING = 0; SKIPPED = 0; REBOOT_REQUIRED = 0
    OFFLINE = 0; SMB_FAIL = 0; NO_CAPABILITY_API = 0
    WINRM_FAIL = 0; PSEXEC_FAIL = 0; FAIL = 0
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

    $state = @{ host = $h.Name; ip = $h.IP; transport = ''; status = 'UNKNOWN'; category = ''; message = ''; exit_code = $null; sshd_status = '' }

    # 1) SSH уже работает?
    if (Test-TcpPort -ComputerName $h.IP -Port $sshPort -TimeoutSec $sshTestTimeout) {
        Write-Log "  [ OK ] $displayName — SSH уже отвечает" "OK"
        $state.status = 'OK'; $state.category = 'ALREADY_WORKING'
        $state.message = "TCP $sshPort already accepting connections"
        $state.transport = 'none'
        Write-HostState -Path $statePath -State $state
        $stats.OK++; $stats.ALREADY_WORKING++
        continue
    }

    $ping = Test-Connection -ComputerName $h.IP -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $ping) {
        Write-Log "  [ OFFLINE ] $displayName" "WARN"
        $state.status = 'OFFLINE'; $state.category = 'OFFLINE'; $state.message = 'ICMP no reply'
        Write-HostState -Path $statePath -State $state
        $stats.OFFLINE++
        continue
    }

    if ($DryRun) {
        $winrmOpen = Test-TcpPort -ComputerName $h.IP -Port 5985 -TimeoutSec 2
        $smbOpen = Test-TcpPort -ComputerName $h.IP -Port 445 -TimeoutSec 2
        $tag = if ($winrmOpen) { 'WinRM' } elseif ($smbOpen) { 'PsExec' } else { 'SMB_FAIL' }
        Write-Log "  [ DRY ] $displayName — пойдёт через $tag (WinRM:$winrmOpen / SMB:$smbOpen)"
        continue
    }

    Write-Log "Установка: $displayName"
    $r = Install-OpenSSH -Target $h -Credential $cred -Username $username -Password $password `
        -SshdSettings $sshdSettings -PsExecPath $psexecPath -MsiPath $msiPath -MsiAvailable $msiAvailable `
        -WinrmTimeout $sshTestTimeout -SmbTimeout $sshTestTimeout

    $state.transport = $r.transport
    $state.status = $r.status
    $state.category = $r.category
    $state.message = $r.message
    if ($r.exit_code) { $state.exit_code = $r.exit_code }
    if ($r.sshd_status) { $state.sshd_status = $r.sshd_status }

    if ($r.status -eq 'OK') {
        # Verify
        Start-Sleep -Seconds 2
        if (Test-TcpPort -ComputerName $h.IP -Port $sshPort -TimeoutSec $sshTestTimeout) {
            Write-Log "  [ OK ] $displayName — SSH работает (transport=$($r.transport))" "OK"
            $stats.OK++
        } else {
            Write-Log "  [ FAIL ] $displayName — установка OK, но SSH не отвечает" "ERROR"
            $state.status = 'FAIL'; $state.category = 'VERIFY_FAIL'
            $state.message = "Install reported OK but TCP $sshPort closed"
            $stats.FAIL++
        }
    } elseif ($r.status -eq 'REBOOT_REQUIRED') {
        Write-Log "  [ REBOOT_REQUIRED ] $displayName" "WARN"
        $stats.REBOOT_REQUIRED++
    } elseif ($r.status -eq 'SMB_FAIL') {
        Write-Log "  [ SMB_FAIL ] $displayName — нет ни WinRM, ни SMB" "WARN"
        $stats.SMB_FAIL++
    } elseif ($r.status -eq 'NO_CAPABILITY_API') {
        Write-Log "  [ NO_CAPABILITY_API ] $displayName — $($r.message)" "WARN"
        $stats.NO_CAPABILITY_API++
    } elseif ($r.transport -like 'winrm*' -and $r.status -ne 'OK') {
        Write-Log "  [ WINRM_FAIL ] $displayName — $($r.message)" "ERROR"
        $stats.WINRM_FAIL++
    } elseif ($r.transport -eq 'psexec') {
        Write-Log "  [ PSEXEC_FAIL ] $displayName — $($r.message)" "ERROR"
        $stats.PSEXEC_FAIL++
    } else {
        Write-Log "  [ FAIL ] $displayName — $($r.message)" "ERROR"
        $stats.FAIL++
    }

    Write-HostState -Path $statePath -State $state
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
Write-Log "Reboot required:     $($stats.REBOOT_REQUIRED)"
Write-Log "Offline:             $($stats.OFFLINE)"
Write-Log "SMB unreachable:     $($stats.SMB_FAIL)"
Write-Log "No Capability API:   $($stats.NO_CAPABILITY_API)"
Write-Log "WinRM failed:        $($stats.WINRM_FAIL)"
Write-Log "PsExec failed:       $($stats.PSEXEC_FAIL)"
Write-Log "Other failed:        $($stats.FAIL)"
Write-Log ""
Write-Log "State: $stateDir"
Write-Log "Лог: $LogFile"

$failures = $stats.WINRM_FAIL + $stats.PSEXEC_FAIL + $stats.FAIL
if ($failures -gt 0) { exit 1 }

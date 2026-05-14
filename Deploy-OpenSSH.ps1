<#
.SYNOPSIS
    Массовая установка и настройка OpenSSH Server на Windows-хостах через PsExec.
.DESCRIPTION
    Скрипт читает список Windows-хостов из inventory.yaml,
    проверяет доступность SSH (порт 22), и если недоступен —
    устанавливает и настраивает OpenSSH Server удалённо через PsExec.

    На этапе установки настраивается вход по паролю (включая доменные учётки).
    Деплой SSH-ключей — отдельный этап.
.PARAMETER DryRun
    Только показать список хостов и статусы, без изменений.
.PARAMETER HostFilter
    Фильтр по имени хоста или IP (поддерживает wildcard *).
.PARAMETER HostsFile
    Путь к файлу со списком хостов (по одному имени/IP на строку,
    '#' — комментарий). Если задан — перекрывает HostFilter.
.PARAMETER ConfigPath
    Путь к config.yaml. По умолчанию — рядом со скриптом.
.EXAMPLE
    .\Deploy-OpenSSH.ps1 -DryRun
    .\Deploy-OpenSSH.ps1 -HostFilter "dev-rds1"
    .\Deploy-OpenSSH.ps1 -HostsFile .\retry-hosts.txt
    .\Deploy-OpenSSH.ps1
#>

param(
    [switch]$DryRun,
    [string]$HostFilter = "*",
    [string]$HostsFile,
    [string]$ConfigPath
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$LogFile = Join-Path $ScriptDir "Deploy-OpenSSH.log"

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
    #>
    param([string]$Path)

    $lines = Get-Content $Path
    $root = @{}
    $stack = @( @{ obj = $root; indent = -1 } )

    foreach ($line in $lines) {
        if ($line -match '^\s*$' -or $line -match '^\s*#') { continue }

        $trimmed = $line.TrimStart()
        $indent = $line.Length - $trimmed.Length

        while ($stack.Count -gt 1 -and $stack[-1].indent -ge $indent) {
            $stack = $stack[0..($stack.Count - 2)]
        }

        $current = $stack[-1].obj

        if ($trimmed -match '^(.+?):\s*(.+)$') {
            $key = $Matches[1].Trim()
            $val = $Matches[2].Trim().Trim('"').Trim("'")
            if ($val -eq '{}') {
                $current[$key] = @{}
            } else {
                $current[$key] = $val
            }
        }
        elseif ($trimmed -match '^(.+?):\s*$') {
            $key = $Matches[1].Trim()
            $newObj = @{}
            $current[$key] = $newObj
            $stack += @{ obj = $newObj; indent = $indent }
        }
    }

    return $root
}

function Find-GroupByName {
    <#
    .SYNOPSIS
        Рекурсивно ищет в дереве inventory группу с указанным именем.
        Возвращает hashtable группы или $null.
    #>
    param([hashtable]$Node, [string]$Name)

    if (-not $Node) { return $null }

    if ($Node.ContainsKey($Name) -and $Node[$Name] -is [hashtable]) {
        return $Node[$Name]
    }

    foreach ($key in $Node.Keys) {
        $val = $Node[$key]
        if ($val -is [hashtable]) {
            if ($val['children'] -is [hashtable]) {
                $found = Find-GroupByName -Node $val['children'] -Name $Name
                if ($found) { return $found }
            }
        }
    }

    return $null
}

function Collect-HostsRecursive {
    <#
    .SYNOPSIS
        Рекурсивно собирает все хосты из группы и её потомков.
    #>
    param([hashtable]$Group, [System.Collections.ArrayList]$Acc)

    if (-not $Group) { return }

    if ($Group['hosts'] -is [hashtable]) {
        foreach ($hostName in $Group['hosts'].Keys) {
            $hostData = $Group['hosts'][$hostName]
            if ($hostData -is [hashtable] -and $hostData['ansible_host']) {
                [void]$Acc.Add(@{
                    Name = $hostName
                    IP   = $hostData['ansible_host']
                })
            }
        }
    }

    if ($Group['children'] -is [hashtable]) {
        foreach ($childName in $Group['children'].Keys) {
            $child = $Group['children'][$childName]
            if ($child -is [hashtable]) {
                Collect-HostsRecursive -Group $child -Acc $Acc
            }
        }
    }
}

function Get-WindowsHosts {
    <#
    .SYNOPSIS
        Извлекает список Windows-хостов из inventory.yaml.
        Возвращает массив @{ Name; IP }.
        Ищет группу 'windows' где угодно в дереве и рекурсивно собирает все хосты.
    #>
    param([hashtable]$Inventory)

    $acc = [System.Collections.ArrayList]::new()

    $root = $Inventory['all']
    if (-not $root) { return @() }

    $windowsGroup = $null
    if ($root['children'] -is [hashtable]) {
        $windowsGroup = Find-GroupByName -Node $root['children'] -Name 'windows'
    }

    if (-not $windowsGroup) { return @() }

    Collect-HostsRecursive -Group $windowsGroup -Acc $acc

    return @($acc)
}

function Test-SSH {
    param([string]$ComputerName, [int]$Port = 22, [int]$TimeoutSec = 5)

    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $connect = $tcp.BeginConnect($ComputerName, $Port, $null, $null)
        $wait = $connect.AsyncWaitHandle.WaitOne($TimeoutSec * 1000, $false)
        if (-not $wait) {
            $tcp.Close()
            return $false
        }
        $tcp.EndConnect($connect)
        $tcp.Close()
        return $true
    }
    catch {
        return $false
    }
}

function Invoke-PsExecOpenSSH {
    <#
    .SYNOPSIS
        Устанавливает и настраивает OpenSSH Server на удалённом хосте через PsExec.
    #>
    param(
        [string]$PsExecPath,
        [string]$ComputerName,
        [string]$Username,
        [string]$Password,
        [hashtable]$SshdSettings
    )

    $sshPort = $SshdSettings['port']
    $pwdAuth = if ($SshdSettings['password_authentication'] -eq 'yes') { 'yes' } else { 'no' }
    $pubkeyAuth = if ($SshdSettings['pubkey_authentication'] -eq 'yes') { 'yes' } else { 'no' }
    $defaultShell = $SshdSettings['default_shell']

    Write-Log "  PsExec -> \\$ComputerName : Install OpenSSH Server"

    # PowerShell-скрипт для выполнения на удалённой машине.
    # Exit codes: 0 = OK; 1 = generic error; 2 = NO_CAPABILITY_API (нужен MSI fallback);
    #             3 = REBOOT_REQUIRED после Add-WindowsCapability.
    $remoteScript = @"
`$ErrorActionPreference = 'Stop'
try {
    if (-not (Get-Command -Name Get-WindowsCapability -ErrorAction SilentlyContinue)) {
        Write-Host 'NO_CAPABILITY_API'
        exit 2
    }
    `$feat = Get-WindowsCapability -Online | Where-Object { `$_.Name -like 'OpenSSH.Server*' }
    if (-not `$feat) {
        Write-Host 'NO_CAPABILITY_API'
        exit 2
    }
    if (`$feat.State -ne 'Installed') {
        Write-Host 'INSTALLING OpenSSH Server...'
        `$res = Add-WindowsCapability -Online -Name 'OpenSSH.Server~~~~0.0.1.0' -ErrorAction Stop
        if (`$res.RestartNeeded) {
            Write-Host 'REBOOT_REQUIRED'
            exit 3
        }
        Write-Host 'INSTALLED'
    } else {
        Write-Host 'ALREADY_INSTALLED'
    }

    `$cfgDir = `$env:ProgramData + '\ssh'
    `$cfgPath = `$cfgDir + '\sshd_config'

    if (-not (Test-Path `$cfgDir)) {
        # Каталога нет — значит Capability "приняли", но файлы не созданы.
        # Это типично для pending-reboot. Просим перезагрузку.
        Write-Host 'REBOOT_REQUIRED'
        exit 3
    }

    if (Test-Path `$cfgPath) {
        `$bak = `$cfgPath + '.bak.' + (Get-Date -Format 'yyyyMMdd_HHmmss')
        Copy-Item `$cfgPath `$bak -ErrorAction Stop
        Write-Host "BACKUP: `$bak"
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
    Write-Host 'SSHD_CONFIG_OK'

    `$regPath = 'HKLM:\SOFTWARE\OpenSSH'
    if (-not (Test-Path `$regPath)) { New-Item -Path `$regPath -Force -ErrorAction Stop | Out-Null }
    Set-ItemProperty -Path `$regPath -Name 'DefaultShell' -Value '$defaultShell' -ErrorAction Stop
    Set-ItemProperty -Path `$regPath -Name 'DefaultShellCommandOption' -Value '/c' -ErrorAction Stop
    Write-Host 'DEFAULT_SHELL_OK'

    `$fw = Get-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -ErrorAction SilentlyContinue
    if (-not `$fw) {
        New-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -DisplayName 'OpenSSH Server (SSH)' -Enabled True -Direction Inbound -Protocol TCP -Action Allow -LocalPort $sshPort -ErrorAction Stop | Out-Null
        Write-Host 'FIREWALL_CREATED'
    } else {
        Enable-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -ErrorAction SilentlyContinue
        Write-Host 'FIREWALL_EXISTS'
    }

    `$svc = Get-Service -Name sshd -ErrorAction Stop
    Set-Service -Name sshd -StartupType Automatic -ErrorAction Stop
    Restart-Service sshd -Force -ErrorAction Stop
    Start-Sleep -Seconds 2
    `$svc = Get-Service -Name sshd -ErrorAction Stop
    Write-Host "SSHD_STATUS: `$(`$svc.Status)"

    if (`$svc.Status -eq 'Running') {
        Write-Host 'OPENSSH_OK'
    } else {
        Write-Host 'OPENSSH_FAIL'
        exit 1
    }
} catch {
    Write-Host "OPENSSH_ERROR: `$(`$_.Exception.Message)"
    exit 1
}
"@

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
        -RedirectStandardOutput "$env:TEMP\psexec_ssh_out.txt" `
        -RedirectStandardError "$env:TEMP\psexec_ssh_err.txt"

    $stdout = if (Test-Path "$env:TEMP\psexec_ssh_out.txt") {
        Get-Content "$env:TEMP\psexec_ssh_out.txt" -Raw
    } else { "" }
    $stderr = if (Test-Path "$env:TEMP\psexec_ssh_err.txt") {
        Get-Content "$env:TEMP\psexec_ssh_err.txt" -Raw
    } else { "" }

    # Выводим детали из stdout
    foreach ($line in ($stdout -split "`n")) {
        $line = $line.Trim()
        if ($line) {
            Write-Log "    $line"
        }
    }

    # Возврат: hashtable с диагностированным типом результата:
    #   Result = 'OK' | 'NEED_MSI' | 'NEED_REBOOT' | 'FAIL'
    #   ExitCode / Stderr — для логов
    $result = @{ ExitCode = $proc.ExitCode; Stderr = $stderr; Result = 'FAIL' }
    if ($proc.ExitCode -eq 0) {
        Write-Log "  Установка завершена успешно" "OK"
        $result.Result = 'OK'
    } elseif ($proc.ExitCode -eq 2) {
        Write-Log "  Нет Windows Capability API — нужен MSI fallback" "WARN"
        $result.Result = 'NEED_MSI'
    } elseif ($proc.ExitCode -eq 3) {
        Write-Log "  Требуется перезагрузка хоста (pending-reboot после установки)" "WARN"
        $result.Result = 'NEED_REBOOT'
    } else {
        Write-Log "  Ошибка (exitcode=$($proc.ExitCode)): $stderr" "ERROR"
    }
    return $result
}

function Invoke-PsExecOpenSshMsi {
    <#
    .SYNOPSIS
        Fallback-установщик через MSI (для старых Windows без Capability API).
        1. Копирует MSI на удалённый хост через \\host\C$\Windows\Temp\.
        2. Запускает msiexec /i ... /qn /norestart через PsExec.
        3. Регистрирует/стартует sshd, пишет sshd_config, DefaultShell, firewall.
    #>
    param(
        [string]$PsExecPath,
        [string]$ComputerName,
        [string]$Username,
        [string]$Password,
        [string]$MsiPath,
        [hashtable]$SshdSettings
    )

    $sshPort = $SshdSettings['port']
    $pwdAuth = if ($SshdSettings['password_authentication'] -eq 'yes') { 'yes' } else { 'no' }
    $pubkeyAuth = if ($SshdSettings['pubkey_authentication'] -eq 'yes') { 'yes' } else { 'no' }
    $defaultShell = $SshdSettings['default_shell']

    Write-Log "  MSI -> \\$ComputerName : копирование $MsiPath"

    $msiName = [System.IO.Path]::GetFileName($MsiPath)
    $remoteUnc = "\\$ComputerName\C`$\Windows\Temp\$msiName"
    $remoteLocal = "C:\Windows\Temp\$msiName"

    # Подключаем admin$ под нужной учёткой (если ещё не подключён).
    # net use молча игнорируем ошибку "уже подключён".
    $netRoot = "\\$ComputerName\C`$"
    & net use $netRoot /user:$Username $Password 2>&1 | Out-Null

    try {
        Copy-Item -Path $MsiPath -Destination $remoteUnc -Force -ErrorAction Stop
    } catch {
        Write-Log "  Не удалось скопировать MSI: $($_.Exception.Message)" "ERROR"
        & net use $netRoot /delete /yes 2>&1 | Out-Null
        return @{ Result = 'FAIL'; ExitCode = -1; Stderr = $_.Exception.Message }
    } finally {
        & net use $netRoot /delete /yes 2>&1 | Out-Null
    }

    Write-Log "  PsExec -> \\$ComputerName : msiexec /i $remoteLocal /qn"

    # Remote script для MSI-варианта: установка + регистрация службы + конфиг.
    $remoteScript = @"
`$ErrorActionPreference = 'Stop'
try {
    `$msiArgs = @('/i', '$remoteLocal', '/qn', '/norestart')
    `$p = Start-Process -FilePath 'msiexec.exe' -ArgumentList `$msiArgs -Wait -PassThru
    if (`$p.ExitCode -ne 0 -and `$p.ExitCode -ne 3010) {
        Write-Host "MSI_INSTALL_FAILED: exitcode `$(`$p.ExitCode)"
        exit 1
    }
    Write-Host 'MSI_INSTALLED'

    `$svc = Get-Service -Name sshd -ErrorAction SilentlyContinue
    if (-not `$svc) {
        `$installScript = 'C:\Program Files\OpenSSH\install-sshd.ps1'
        if (Test-Path `$installScript) {
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File `$installScript | Out-Null
            Write-Host 'SSHD_REGISTERED'
        } else {
            Write-Host 'INSTALL_SCRIPT_NOT_FOUND'
            exit 1
        }
        `$svc = Get-Service -Name sshd -ErrorAction Stop
    } else {
        Write-Host 'SSHD_ALREADY_REGISTERED'
    }

    Set-Service -Name sshd -StartupType Automatic -ErrorAction Stop
    if (`$svc.Status -ne 'Running') {
        Start-Service sshd -ErrorAction Stop
        Start-Sleep -Seconds 2
    }

    `$cfgDir = `$env:ProgramData + '\ssh'
    `$cfgPath = `$cfgDir + '\sshd_config'
    if (-not (Test-Path `$cfgDir)) {
        New-Item -ItemType Directory -Path `$cfgDir -Force -ErrorAction Stop | Out-Null
    }
    if (Test-Path `$cfgPath) {
        `$bak = `$cfgPath + '.bak.' + (Get-Date -Format 'yyyyMMdd_HHmmss')
        Copy-Item `$cfgPath `$bak -ErrorAction Stop
        Write-Host "BACKUP: `$bak"
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
    Write-Host 'SSHD_CONFIG_OK'

    `$regPath = 'HKLM:\SOFTWARE\OpenSSH'
    if (-not (Test-Path `$regPath)) { New-Item -Path `$regPath -Force -ErrorAction Stop | Out-Null }
    Set-ItemProperty -Path `$regPath -Name 'DefaultShell' -Value '$defaultShell' -ErrorAction Stop
    Set-ItemProperty -Path `$regPath -Name 'DefaultShellCommandOption' -Value '/c' -ErrorAction Stop
    Write-Host 'DEFAULT_SHELL_OK'

    `$fw = Get-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -ErrorAction SilentlyContinue
    if (-not `$fw) {
        # На Win 7 / 2008 Get-NetFirewallRule может отсутствовать —
        # тогда добавим правило через netsh.
        if (Get-Command -Name New-NetFirewallRule -ErrorAction SilentlyContinue) {
            New-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -DisplayName 'OpenSSH Server (SSH)' -Enabled True -Direction Inbound -Protocol TCP -Action Allow -LocalPort $sshPort -ErrorAction Stop | Out-Null
        } else {
            & netsh.exe advfirewall firewall add rule name="OpenSSH-Server-In-TCP" dir=in action=allow protocol=TCP localport=$sshPort | Out-Null
        }
        Write-Host 'FIREWALL_CREATED'
    } else {
        Enable-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -ErrorAction SilentlyContinue
        Write-Host 'FIREWALL_EXISTS'
    }

    Restart-Service sshd -Force -ErrorAction Stop
    Start-Sleep -Seconds 2
    `$svc = Get-Service -Name sshd -ErrorAction Stop
    Write-Host "SSHD_STATUS: `$(`$svc.Status)"

    if (`$svc.Status -eq 'Running') {
        Write-Host 'OPENSSH_OK'
    } else {
        Write-Host 'OPENSSH_FAIL'
        exit 1
    }
} catch {
    Write-Host "OPENSSH_ERROR: `$(`$_.Exception.Message)"
    exit 1
} finally {
    Remove-Item -Path '$remoteLocal' -Force -ErrorAction SilentlyContinue
}
"@

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
        -RedirectStandardOutput "$env:TEMP\psexec_msi_out.txt" `
        -RedirectStandardError "$env:TEMP\psexec_msi_err.txt"

    $stdout = if (Test-Path "$env:TEMP\psexec_msi_out.txt") {
        Get-Content "$env:TEMP\psexec_msi_out.txt" -Raw
    } else { "" }
    $stderr = if (Test-Path "$env:TEMP\psexec_msi_err.txt") {
        Get-Content "$env:TEMP\psexec_msi_err.txt" -Raw
    } else { "" }

    foreach ($line in ($stdout -split "`n")) {
        $line = $line.Trim()
        if ($line) { Write-Log "    $line" }
    }

    $result = @{ ExitCode = $proc.ExitCode; Stderr = $stderr; Result = 'FAIL' }
    if ($proc.ExitCode -eq 0) {
        Write-Log "  MSI-установка завершена успешно" "OK"
        $result.Result = 'OK'
    } else {
        Write-Log "  Ошибка MSI (exitcode=$($proc.ExitCode)): $stderr" "ERROR"
    }
    return $result
}

# ============================================================
# MAIN
# ============================================================

Write-Log "=========================================="
Write-Log "=== Deploy-OpenSSH — Начало ==="
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
$sshTimeout = [int]($config['ssh_test_timeout'])
$msiPath = $config['openssh_msi_path']

# Настройки sshd
$sshdSettings = @{
    port                     = if ($config['sshd'] -and $config['sshd']['port']) { $config['sshd']['port'] } else { "22" }
    password_authentication  = if ($config['sshd'] -and $config['sshd']['password_authentication']) { $config['sshd']['password_authentication'] } else { "yes" }
    pubkey_authentication    = if ($config['sshd'] -and $config['sshd']['pubkey_authentication']) { $config['sshd']['pubkey_authentication'] } else { "yes" }
    default_shell            = if ($config['sshd'] -and $config['sshd']['default_shell']) { $config['sshd']['default_shell'] } else { "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" }
}

$sshPort = [int]$sshdSettings['port']

if (-not (Test-Path $psexecPath)) {
    Write-Log "PsExec не найден: $psexecPath" "ERROR"
    exit 1
}

if (-not (Test-Path $inventoryPath)) {
    Write-Log "Inventory не найден: $inventoryPath" "ERROR"
    exit 1
}

$msiAvailable = $false
if ($msiPath) {
    if (Test-Path $msiPath) {
        $msiAvailable = $true
        Write-Log "MSI fallback доступен: $msiPath"
    } else {
        Write-Log "openssh_msi_path задан, но файл не найден: $msiPath — MSI fallback отключён" "WARN"
    }
} else {
    Write-Log "openssh_msi_path не задан — MSI fallback отключён" "WARN"
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

# Фильтрация: HostsFile (если задан) приоритетнее HostFilter.
if ($HostsFile) {
    if (-not (Test-Path $HostsFile)) {
        Write-Log "HostsFile не найден: $HostsFile" "ERROR"
        exit 1
    }
    $wanted = @{}
    foreach ($line in Get-Content $HostsFile) {
        $line = $line.Trim()
        if (-not $line -or $line.StartsWith('#')) { continue }
        $wanted[$line.ToLower()] = $true
    }
    Write-Log "HostsFile: $HostsFile (записей: $($wanted.Count))"
    $filteredHosts = $allHosts | Where-Object {
        $wanted.ContainsKey($_.Name.ToLower()) -or $wanted.ContainsKey($_.IP.ToLower())
    }
} else {
    $filteredHosts = $allHosts | Where-Object {
        $_.Name -like $HostFilter -or $_.IP -like $HostFilter
    }
}

# Сортировка по IP
$filteredHosts = $filteredHosts | Sort-Object { [version]($_.IP -replace '(\d+)\.(\d+)\.(\d+)\.(\d+)', '$1.$2.$3.$4') }

Write-Log "Всего Windows-хостов: $($allHosts.Count)"
if ($HostsFile) {
    Write-Log "После HostsFile: $($filteredHosts.Count)"
} else {
    Write-Log "После фильтра '$HostFilter': $($filteredHosts.Count)"
}

if ($filteredHosts.Count -eq 0) {
    Write-Log "Нет хостов для обработки" "WARN"
    exit 0
}

# --- 4. Фаза 1: Проверка SSH ---
Write-Log ""
Write-Log "=== Фаза 1: Проверка SSH ==="
Write-Log ""

$sshOK = @()
$sshFail = @()
$unreachable = @()

foreach ($h in $filteredHosts) {
    $displayName = "$($h.Name) ($($h.IP))"

    $ping = Test-Connection -ComputerName $h.IP -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $ping) {
        Write-Log "  [ OFFLINE ] $displayName" "WARN"
        $unreachable += $h
        continue
    }

    $sshStatus = Test-SSH -ComputerName $h.IP -Port $sshPort -TimeoutSec $sshTimeout
    if ($sshStatus) {
        Write-Log "  [  OK  ] $displayName — SSH работает (порт $sshPort)" "OK"
        $sshOK += $h
    } else {
        Write-Log "  [ FAIL ] $displayName — SSH недоступен"
        $sshFail += $h
    }
}

Write-Log ""
Write-Log "Итого фазы 1: OK=$($sshOK.Count)  FAIL=$($sshFail.Count)  OFFLINE=$($unreachable.Count)"

if ($DryRun) {
    Write-Log ""
    Write-Log "=== DryRun — без изменений ==="
    Write-Log "Хосты для установки OpenSSH:"
    foreach ($h in $sshFail) {
        Write-Log "  - $($h.Name) ($($h.IP))"
    }
    exit 0
}

if ($sshFail.Count -eq 0) {
    Write-Log "Все хосты уже с SSH — установка не нужна" "OK"
    exit 0
}

# --- 5. Фаза 2: Установка OpenSSH через PsExec ---
Write-Log ""
Write-Log "=== Фаза 2: Установка OpenSSH через PsExec ==="
Write-Log ""

$configuredOK = @()
$configuredFail = @()
$needMsi = @()
$needReboot = @()

foreach ($h in $sshFail) {
    $displayName = "$($h.Name) ($($h.IP))"
    Write-Log "Установка: $displayName"

    try {
        $r = Invoke-PsExecOpenSSH `
            -PsExecPath $psexecPath `
            -ComputerName $h.IP `
            -Username $username `
            -Password $password `
            -SshdSettings $sshdSettings

        switch ($r.Result) {
            'OK'           { $configuredOK += $h }
            'NEED_MSI'     { $needMsi += $h }
            'NEED_REBOOT'  { $needReboot += $h }
            default        { $configuredFail += $h }
        }
    }
    catch {
        Write-Log "  Исключение: $_" "ERROR"
        $configuredFail += $h
    }
}

# --- 5b. Фаза 2b: MSI fallback для хостов без Capability API ---
if ($needMsi.Count -gt 0) {
    Write-Log ""
    Write-Log "=== Фаза 2b: MSI fallback для $($needMsi.Count) хост(ов) ==="
    Write-Log ""

    if (-not $msiAvailable) {
        Write-Log "MSI недоступен — пропускаем фоллбек. Эти хосты пойдут в FAIL." "WARN"
        $configuredFail += $needMsi
    } else {
        foreach ($h in $needMsi) {
            $displayName = "$($h.Name) ($($h.IP))"
            Write-Log "MSI-установка: $displayName"

            try {
                $r = Invoke-PsExecOpenSshMsi `
                    -PsExecPath $psexecPath `
                    -ComputerName $h.IP `
                    -Username $username `
                    -Password $password `
                    -MsiPath $msiPath `
                    -SshdSettings $sshdSettings

                if ($r.Result -eq 'OK') {
                    $configuredOK += $h
                } else {
                    $configuredFail += $h
                }
            }
            catch {
                Write-Log "  Исключение: $_" "ERROR"
                $configuredFail += $h
            }
        }
    }
}

# --- 6. Фаза 3: Повторная проверка SSH ---
Write-Log ""
Write-Log "=== Фаза 3: Повторная проверка SSH ==="
Write-Log ""

Start-Sleep -Seconds 3

$verifyOK = @()
$verifyFail = @()

foreach ($h in $configuredOK) {
    $displayName = "$($h.Name) ($($h.IP))"
    $sshStatus = Test-SSH -ComputerName $h.IP -Port $sshPort -TimeoutSec $sshTimeout

    if ($sshStatus) {
        Write-Log "  [  OK  ] $displayName — SSH работает" "OK"
        $verifyOK += $h
    } else {
        Write-Log "  [ FAIL ] $displayName — SSH всё ещё недоступен" "ERROR"
        $verifyFail += $h
    }
}

# --- 7. Итоговый отчёт ---
Write-Log ""
Write-Log "=========================================="
Write-Log "=== ИТОГО ==="
Write-Log "=========================================="
Write-Log "Всего хостов:           $($filteredHosts.Count)"
Write-Log "SSH уже работал:        $($sshOK.Count)"
Write-Log "Офлайн:                 $($unreachable.Count)"
Write-Log "Установлено успешно:    $($verifyOK.Count)"
Write-Log "Установлено с ошибками: $($configuredFail.Count)"
Write-Log "Не прошло проверку:     $($verifyFail.Count)"
Write-Log "Требует перезагрузки:   $($needReboot.Count)"

if ($unreachable.Count -gt 0) {
    Write-Log ""
    Write-Log "Офлайн хосты:"
    foreach ($h in $unreachable) {
        Write-Log "  - $($h.Name) ($($h.IP))" "WARN"
    }
}

if ($needReboot.Count -gt 0) {
    Write-Log ""
    Write-Log "Хосты, требующие перезагрузки (Add-WindowsCapability принят, но pending-reboot):"
    foreach ($h in $needReboot) {
        Write-Log "  - $($h.Name) ($($h.IP)) — после ребута запустить скрипт ещё раз" "WARN"
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

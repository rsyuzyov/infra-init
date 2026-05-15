<#
.SYNOPSIS
    Массовое прописывание SSH публичного ключа на Windows-хостах.
.DESCRIPTION
    Транспорты по приоритету для каждого хоста:
      1. SSH (Posh-SSH, парольная авторизация) — если TCP 22 открыт и пароль подходит.
         После добавления — post-check входа по приватному ключу.
      2. WinRM (Invoke-Command) — fallback:
         - если TCP 22 закрыт (кладём ключ авансом, на будущее)
         - если SSH-сервер требует pubkey-only (паролем не зайти)
         - если что-то ещё помешало SSH-варианту

    Per-host state: output/<domain>/ssh-key/<host>.json.
    Уже OK-хосты пропускаются; -Force переобрабатывает.

.PARAMETER DryRun
    Только показать план.
.PARAMETER HostFilter
    Wildcard-фильтр по имени или IP.
.PARAMETER HostsFile
    Файл со списком хостов (приоритетнее HostFilter).
.PARAMETER ConfigPath
    Путь к config.yaml.
.PARAMETER PublicKeyPath
    Переопределить public_key_path.
.PARAMETER Force
    Игнорировать сохранённый state.
.PARAMETER Remove
    Удалить ключ (по совпадению keyBody) из administrators_authorized_keys и ~/.ssh/authorized_keys.
    Игнорирует сохранённый state (как -Force).

.EXAMPLE
    .\Deploy-SshKey.ps1 -DryRun
    .\Deploy-SshKey.ps1 -HostsFile .\tmp\hosts.txt
    .\Deploy-SshKey.ps1 -Force
    .\Deploy-SshKey.ps1 -Remove
#>

param(
    [switch]$DryRun,
    [string]$HostFilter = "*",
    [string]$HostsFile,
    [string]$ConfigPath,
    [string]$PublicKeyPath,
    [switch]$Force,
    [switch]$Remove
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $ScriptDir "Common.ps1")

# Админ нужен только чтобы донастроить локальный WSMan-клиент (TrustedHosts, Basic, AllowUnencrypted).
# Если он уже настроен — повышение не запрашиваем.
Assert-Admin -ScriptPath $MyInvocation.MyCommand.Definition -BoundParams $PSBoundParameters `
    -SkipIf { Test-WSManClientReady } -Reason "настройка локального WSMan-клиента"

$LogFile = New-LogFile -RootDir $ScriptDir -ScriptName "Deploy-SshKey"

# ============================================================
# Удалённый скрипт (общий для SSH и WinRM транспортов).
# Возвращает hashtable (для WinRM) или печатает JSON в маркерах (для SSH-encoded).
# Параметры — через ArgumentList в Invoke-Command, через .Replace() для encoded-варианта.
# ============================================================

$SshKeyDeployScript = {
    param($publicKey, $targetUser, $removeMode)
    $ErrorActionPreference = 'Stop'
    $result = @{ status = 'ERROR'; message = ''; file = ''; is_admin = $false; acl_err = $null; removed_from = @() }
    try {
        $acct = New-Object System.Security.Principal.NTAccount($targetUser)
        $userSid = $acct.Translate([System.Security.Principal.SecurityIdentifier]).Value

        # Детектим админство. Если targetUser совпадает с текущим залогиненным —
        # используем access token (учитывает вложенность Domain Admins → BUILTIN\Administrators).
        # Иначе — перечисление прямых членов локальной группы (есть ограничение: вложенные группы не покроются).
        $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $currentNT = $currentIdentity.Name
        $currentShort = $currentNT.Split('\')[-1]
        $targetShort = if ($targetUser -match '\\') { $targetUser.Split('\')[-1] } else { $targetUser }
        $sameUser = ($currentNT -ieq $targetUser) -or
                    (-not ($targetUser -match '\\') -and ($currentShort -ieq $targetShort))

        $isAdmin = $false
        if ($sameUser) {
            $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
            $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        } else {
            try {
                $members = Get-LocalGroupMember -SID 'S-1-5-32-544' -ErrorAction Stop
                foreach ($m in $members) {
                    if ($m.SID.Value -eq $userSid) { $isAdmin = $true; break }
                }
            } catch {
                $grpName = (New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')).Translate([System.Security.Principal.NTAccount]).Value
                $shortGroup = $grpName.Split('\')[-1]
                $group = [ADSI]"WinNT://./$shortGroup,group"
                foreach ($m in @($group.psbase.Invoke('Members'))) {
                    $sidBytes = $m.GetType().InvokeMember('objectSid', 'GetProperty', $null, $m, $null)
                    $sid = (New-Object System.Security.Principal.SecurityIdentifier($sidBytes, 0)).Value
                    if ($sid -eq $userSid) { $isAdmin = $true; break }
                }
            }
        }

        $shortName = $targetShort

        $adminFile = 'C:\ProgramData\ssh\administrators_authorized_keys'
        $userProfile = "C:\Users\$shortName"
        $userFile = if (Test-Path $userProfile) { Join-Path $userProfile '.ssh\authorized_keys' } else { $null }

        $keyParts = $publicKey.Trim() -split '\s+'
        if ($keyParts.Count -lt 2) { throw 'Invalid public key format' }
        $keyBody = $keyParts[1]

        $result.is_admin = $isAdmin

        if ($removeMode) {
            $removedFrom = @()
            foreach ($f in @($adminFile, $userFile)) {
                if (-not $f) { continue }
                if (-not (Test-Path $f)) { continue }
                # На случай битого DACL у admin-файла берём ownership заранее
                if ($isAdmin -and $f -eq $adminFile) {
                    & takeown.exe /F $f /A *>$null
                    & icacls.exe $f /grant "*S-1-5-32-544:F" *>$null
                }
                $orig = @(Get-Content $f -ErrorAction SilentlyContinue)
                $kept = New-Object System.Collections.ArrayList
                $found = $false
                foreach ($l in $orig) {
                    $trim = $l.Trim()
                    if ($trim -and -not $trim.StartsWith('#')) {
                        $parts = $trim -split '\s+'
                        if ($parts.Count -ge 2 -and $parts[1] -eq $keyBody) { $found = $true; continue }
                    }
                    [void]$kept.Add($l)
                }
                if ($found) {
                    Set-Content -Path $f -Value $kept -Encoding ASCII -ErrorAction Stop
                    $removedFrom += $f
                }
            }
            $result.removed_from = $removedFrom
            $result.file = ($removedFrom -join '; ')
            if ($removedFrom.Count -gt 0) { $result.status = 'REMOVED' } else { $result.status = 'NOT_FOUND' }
            return $result
        }

        if ($isAdmin) {
            $sshDir = 'C:\ProgramData\ssh'
            $keyFile = $adminFile
        } else {
            if (-not (Test-Path $userProfile)) {
                throw "Profile not found: $userProfile (non-admin user needs existing profile)"
            }
            $sshDir = Join-Path $userProfile '.ssh'
            $keyFile = $userFile
        }
        $result.file = $keyFile

        if (-not (Test-Path $sshDir)) {
            New-Item -ItemType Directory -Path $sshDir -Force | Out-Null
        }

        # Если admin-файл уже существует — на случай битого DACL от прошлых запусков
        # берём ownership на Administrators и временно даём им :F, чтобы можно было читать/писать.
        # Финальный ACL (inheritance:r + только Admin/SYSTEM) применяется ниже.
        if ($isAdmin -and (Test-Path -LiteralPath $keyFile)) {
            & takeown.exe /F $keyFile /A *>$null
            & icacls.exe $keyFile /grant "*S-1-5-32-544:F" *>$null
        }

        $alreadyExists = $false
        if (Test-Path $keyFile) {
            foreach ($l in (Get-Content $keyFile -ErrorAction SilentlyContinue)) {
                $l = $l.Trim()
                if (-not $l -or $l.StartsWith('#')) { continue }
                $parts = $l -split '\s+'
                if ($parts.Count -ge 2 -and $parts[1] -eq $keyBody) { $alreadyExists = $true; break }
            }
        }

        if ($alreadyExists) {
            $result.status = 'SKIP_EXISTS'
        } else {
            Add-Content -Path $keyFile -Value $publicKey.Trim() -Encoding ASCII
            $result.status = 'ADDED'
        }

        # ACL: через SID (язык-нейтрально, не зависит от резолва имён на RU-Windows).
        # S-1-5-32-544 = BUILTIN\Administrators, S-1-5-18 = NT AUTHORITY\SYSTEM.
        $aclErrs = @()
        $null = & icacls $keyFile /inheritance:r 2>$null
        if ($LASTEXITCODE -ne 0) { $aclErrs += "inheritance:r exit=$LASTEXITCODE" }
        if ($isAdmin) {
            $null = & icacls $keyFile /grant "*S-1-5-32-544:F" "*S-1-5-18:F" 2>$null
        } else {
            $null = & icacls $keyFile /grant "*${userSid}:F" "*S-1-5-18:F" 2>$null
        }
        if ($LASTEXITCODE -ne 0) { $aclErrs += "/grant exit=$LASTEXITCODE" }
        if ($aclErrs.Count -gt 0) { $result.acl_err = ($aclErrs -join '; ') }
    } catch {
        $result.status = 'ERROR'
        $result.message = $_.Exception.Message
    }
    return $result
}

# Для SSH-encoded варианта — то же тело + JSON в маркерах в stdout.
function Build-SshEncodedCommand {
    param([string]$PublicKey, [string]$TargetUser, [bool]$RemoveMode)
    $pkEsc = $PublicKey -replace "'", "''"
    $tuEsc = $TargetUser -replace "'", "''"
    $rmFlag = if ($RemoveMode) { '$true' } else { '$false' }

    $template = @'
$ErrorActionPreference = 'Stop'
$result = @{ status = 'ERROR'; message = ''; file = ''; is_admin = $false; acl_err = $null; removed_from = @() }
try {
    $publicKey  = '__PUBKEY__'
    $targetUser = '__TARGETUSER__'
    $removeMode = __REMOVE__
    $acct = New-Object System.Security.Principal.NTAccount($targetUser)
    $userSid = $acct.Translate([System.Security.Principal.SecurityIdentifier]).Value
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $currentNT = $currentIdentity.Name
    $currentShort = $currentNT.Split('\')[-1]
    $targetShort = if ($targetUser -match '\\') { $targetUser.Split('\')[-1] } else { $targetUser }
    $sameUser = ($currentNT -ieq $targetUser) -or (-not ($targetUser -match '\\') -and ($currentShort -ieq $targetShort))
    $isAdmin = $false
    if ($sameUser) {
        $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
        $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } else {
        try {
            $members = Get-LocalGroupMember -SID 'S-1-5-32-544' -ErrorAction Stop
            foreach ($m in $members) { if ($m.SID.Value -eq $userSid) { $isAdmin = $true; break } }
        } catch {
            $grpName = (New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')).Translate([System.Security.Principal.NTAccount]).Value
            $shortGroup = $grpName.Split('\')[-1]
            $group = [ADSI]"WinNT://./$shortGroup,group"
            foreach ($m in @($group.psbase.Invoke('Members'))) {
                $sidBytes = $m.GetType().InvokeMember('objectSid', 'GetProperty', $null, $m, $null)
                $sid = (New-Object System.Security.Principal.SecurityIdentifier($sidBytes, 0)).Value
                if ($sid -eq $userSid) { $isAdmin = $true; break }
            }
        }
    }
    $shortName = $targetShort
    $adminFile = 'C:\ProgramData\ssh\administrators_authorized_keys'
    $userProfile = "C:\Users\$shortName"
    $userFile = if (Test-Path $userProfile) { Join-Path $userProfile '.ssh\authorized_keys' } else { $null }
    $keyParts = $publicKey.Trim() -split '\s+'
    if ($keyParts.Count -lt 2) { throw 'Invalid public key format' }
    $keyBody = $keyParts[1]
    $result.is_admin = $isAdmin
    if ($removeMode) {
        $removedFrom = @()
        foreach ($f in @($adminFile, $userFile)) {
            if (-not $f) { continue }
            if (-not (Test-Path $f)) { continue }
            if ($isAdmin -and $f -eq $adminFile) {
                & takeown.exe /F $f /A *>$null
                & icacls.exe $f /grant "*S-1-5-32-544:F" *>$null
            }
            $orig = @(Get-Content $f -ErrorAction SilentlyContinue)
            $kept = New-Object System.Collections.ArrayList
            $found = $false
            foreach ($l in $orig) {
                $trim = $l.Trim()
                if ($trim -and -not $trim.StartsWith('#')) {
                    $parts = $trim -split '\s+'
                    if ($parts.Count -ge 2 -and $parts[1] -eq $keyBody) { $found = $true; continue }
                }
                [void]$kept.Add($l)
            }
            if ($found) {
                Set-Content -Path $f -Value $kept -Encoding ASCII -ErrorAction Stop
                $removedFrom += $f
            }
        }
        $result.removed_from = $removedFrom
        $result.file = ($removedFrom -join '; ')
        if ($removedFrom.Count -gt 0) { $result.status = 'REMOVED' } else { $result.status = 'NOT_FOUND' }
    } else {
        if ($isAdmin) {
            $sshDir = 'C:\ProgramData\ssh'
            $keyFile = $adminFile
        } else {
            if (-not (Test-Path $userProfile)) { throw "Profile not found: $userProfile" }
            $sshDir = Join-Path $userProfile '.ssh'
            $keyFile = $userFile
        }
        $result.file = $keyFile
        if (-not (Test-Path $sshDir)) { New-Item -ItemType Directory -Path $sshDir -Force | Out-Null }
        if ($isAdmin -and (Test-Path -LiteralPath $keyFile)) {
            & takeown.exe /F $keyFile /A *>$null
            & icacls.exe $keyFile /grant "*S-1-5-32-544:F" *>$null
        }
        $alreadyExists = $false
        if (Test-Path $keyFile) {
            foreach ($l in (Get-Content $keyFile -ErrorAction SilentlyContinue)) {
                $l = $l.Trim()
                if (-not $l -or $l.StartsWith('#')) { continue }
                $parts = $l -split '\s+'
                if ($parts.Count -ge 2 -and $parts[1] -eq $keyBody) { $alreadyExists = $true; break }
            }
        }
        if ($alreadyExists) { $result.status = 'SKIP_EXISTS' }
        else { Add-Content -Path $keyFile -Value $publicKey.Trim() -Encoding ASCII; $result.status = 'ADDED' }
        $aclErrs = @()
        $null = & icacls $keyFile /inheritance:r 2>$null
        if ($LASTEXITCODE -ne 0) { $aclErrs += "inheritance:r exit=$LASTEXITCODE" }
        if ($isAdmin) {
            $null = & icacls $keyFile /grant "*S-1-5-32-544:F" "*S-1-5-18:F" 2>$null
        } else {
            $null = & icacls $keyFile /grant "*${userSid}:F" "*S-1-5-18:F" 2>$null
        }
        if ($LASTEXITCODE -ne 0) { $aclErrs += "/grant exit=$LASTEXITCODE" }
        if ($aclErrs.Count -gt 0) { $result.acl_err = ($aclErrs -join '; ') }
    }
} catch { $result.status = 'ERROR'; $result.message = $_.Exception.Message }
'__SSHKEYDEPLOY_BEGIN__' + ($result | ConvertTo-Json -Compress) + '__SSHKEYDEPLOY_END__'
'@
    $script = $template.Replace('__PUBKEY__', $pkEsc).Replace('__TARGETUSER__', $tuEsc).Replace('__REMOVE__', $rmFlag)
    $bytes = [System.Text.Encoding]::Unicode.GetBytes($script)
    $b64 = [Convert]::ToBase64String($bytes)
    return "powershell.exe -NoProfile -ExecutionPolicy Bypass -EncodedCommand $b64"
}

function Parse-SshEncodedResult {
    param([string]$Output)
    $m = [regex]::Match($Output, '__SSHKEYDEPLOY_BEGIN__(?<json>.*?)__SSHKEYDEPLOY_END__', 'Singleline')
    if (-not $m.Success) { return $null }
    try { return $m.Groups['json'].Value | ConvertFrom-Json } catch { return $null }
}

# ============================================================
# SSH-транспорт (Posh-SSH)
# ============================================================

function Invoke-SshKeyDeploy {
    param(
        [string]$ComputerName,
        [int]$Port,
        [PSCredential]$Credential,
        [string]$PublicKey,
        [string]$TargetUser,
        [int]$TimeoutSec,
        [bool]$RemoveMode = $false
    )
    $session = $null
    try {
        $session = New-SSHSession -ComputerName $ComputerName -Port $Port -Credential $Credential `
            -AcceptKey -ConnectionTimeout $TimeoutSec -ErrorAction Stop
    } catch {
        return @{ status = 'SSH_AUTH_FAIL'; message = $_.Exception.Message }
    }
    try {
        $cmd = Build-SshEncodedCommand -PublicKey $PublicKey -TargetUser $TargetUser -RemoveMode $RemoveMode
        $cr = Invoke-SSHCommand -SessionId $session.SessionId -Command $cmd -TimeOut 60 -ErrorAction Stop
        $parsed = Parse-SshEncodedResult -Output ($cr.Output -join "`n")
        if (-not $parsed) {
            return @{ status = 'ERROR'; message = "Cannot parse output: $(($cr.Output -join ' | '))" }
        }
        return @{
            status       = $parsed.status
            message      = $parsed.message
            file         = $parsed.file
            is_admin     = $parsed.is_admin
            acl_err      = $parsed.acl_err
            removed_from = $parsed.removed_from
        }
    } catch {
        return @{ status = 'SSH_EXEC_FAIL'; message = $_.Exception.Message }
    } finally {
        if ($session) { Remove-SSHSession -SessionId $session.SessionId -ErrorAction SilentlyContinue | Out-Null }
    }
}

function Test-SshKeyLogin {
    # Проверяет, что вход по приватному ключу работает (для post-check).
    param(
        [string]$ComputerName, [int]$Port, [string]$Username,
        [string]$PrivateKeyPath, [int]$TimeoutSec
    )
    $dummy = ConvertTo-SecureString "no-such-password-key-only" -AsPlainText -Force
    $keyCred = New-Object System.Management.Automation.PSCredential($Username, $dummy)
    $session = $null
    try {
        $session = New-SSHSession -ComputerName $ComputerName -Port $Port -Credential $keyCred `
            -KeyFile $PrivateKeyPath -AcceptKey -ConnectionTimeout $TimeoutSec -ErrorAction Stop
        return @{ Ok = $true; Error = '' }
    } catch {
        return @{ Ok = $false; Error = $_.Exception.Message }
    } finally {
        if ($session) { Remove-SSHSession -SessionId $session.SessionId -ErrorAction SilentlyContinue | Out-Null }
    }
}

# ============================================================
# WinRM-транспорт (fallback)
# ============================================================

function Invoke-WinRmKeyDeploy {
    param(
        [string]$ComputerName,
        [PSCredential]$Credential,
        [string]$PublicKey,
        [string]$TargetUser,
        [string]$AuthMethod = 'Default',
        [bool]$RemoveMode = $false
    )
    $session = $null
    try {
        $session = New-PSSession -ComputerName $ComputerName -Credential $Credential `
            -Authentication $AuthMethod -ErrorAction Stop
    } catch {
        return @{ status = 'WINRM_CONNECT_FAIL'; message = $_.Exception.Message }
    }
    try {
        $r = Invoke-Command -Session $session -ScriptBlock $SshKeyDeployScript `
            -ArgumentList @($PublicKey, $TargetUser, $RemoveMode) -ErrorAction Stop
        return @{
            status       = $r.status
            message      = $r.message
            file         = $r.file
            is_admin     = $r.is_admin
            acl_err      = $r.acl_err
            removed_from = $r.removed_from
        }
    } catch {
        return @{ status = 'WINRM_INVOKE_FAIL'; message = $_.Exception.Message }
    } finally {
        if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
    }
}

# ============================================================
# MAIN
# ============================================================

Write-Log "=========================================="
Write-Log "=== Deploy-SshKey — Начало ==="
Write-Log "=========================================="
Write-Log "Лог: $LogFile"

# Posh-SSH — нужен только для боевого запуска (SSH-транспорт).
# В DryRun проверяем но не падаем, чтобы можно было увидеть план.
$poshSshAvailable = [bool](Get-Module -ListAvailable -Name Posh-SSH)
if (-not $poshSshAvailable) {
    if ($DryRun) {
        Write-Log "Модуль Posh-SSH не установлен — SSH-транспорт будет недоступен (DryRun продолжаем)" "WARN"
    } else {
        Write-Log "Модуль Posh-SSH не установлен." "ERROR"
        Write-Log "Установи: Install-Module -Name Posh-SSH -Scope CurrentUser -Force" "ERROR"
        exit 1
    }
} else {
    Import-Module Posh-SSH -ErrorAction Stop | Out-Null
}

# Конфиг
if (-not $ConfigPath) { $ConfigPath = Join-Path $ScriptDir "config.yaml" }
if (-not (Test-Path $ConfigPath)) {
    Write-Log "Конфиг не найден: $ConfigPath" "ERROR"; exit 1
}
Write-Log "Конфиг: $ConfigPath"
$config = Parse-SimpleYaml -Path $ConfigPath

function _AsString($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [hashtable]) { return $null }
    return [string]$v
}

$inventoryPath = _AsString $config['inventory_path']
$username      = _AsString $config['credentials']['username']
$targetUser    = _AsString $config['target_user']
$pubKeyPath    = if ($PublicKeyPath) { $PublicKeyPath } else { _AsString $config['public_key_path'] }
$sshPort       = if ($config['ssh_port'] -and -not ($config['ssh_port'] -is [hashtable])) { [int]$config['ssh_port'] } else { 22 }
$sshTimeout    = if ($config['ssh_timeout'] -and -not ($config['ssh_timeout'] -is [hashtable])) { [int]$config['ssh_timeout'] } else { 10 }
$domain        = if ($config['domain'] -and -not ($config['domain'] -is [hashtable])) { [string]$config['domain'] } else { '_default' }

if (-not $targetUser) { $targetUser = $username }

if (-not $pubKeyPath -or -not (Test-Path $pubKeyPath)) {
    Write-Log "Публичный ключ не найден: $pubKeyPath" "ERROR"; exit 1
}
if (-not (Test-Path $inventoryPath)) {
    Write-Log "Inventory не найден: $inventoryPath" "ERROR"; exit 1
}

$privateKeyPath = _AsString $config['private_key_path']
if (-not $privateKeyPath -and $pubKeyPath.ToLower().EndsWith('.pub')) {
    $privateKeyPath = $pubKeyPath.Substring(0, $pubKeyPath.Length - 4)
}
$postCheckEnabled = $privateKeyPath -and (Test-Path $privateKeyPath)
if ($privateKeyPath -and -not $postCheckEnabled) {
    Write-Log "Приватный ключ не найден: $privateKeyPath — post-check отключён" "WARN"
}

$publicKey = (Get-Content $pubKeyPath -Raw).Trim()
$keyParts = $publicKey -split '\s+'
if (-not $publicKey -or $keyParts.Count -lt 2) {
    Write-Log "Файл ключа не похож на authorized_keys: $pubKeyPath" "ERROR"; exit 1
}
$fp = $keyParts[1].Substring(0, [Math]::Min(16, $keyParts[1].Length)) + "..."
Write-Log "Публичный ключ: $pubKeyPath ($fp)"
if ($postCheckEnabled) { Write-Log "Приватный ключ для post-check: $privateKeyPath" }
Write-Log "SSH login: $username, target user on hosts: $targetUser"
Write-Log "SSH port: $sshPort, timeout: ${sshTimeout}s"
Write-Log "Domain: $domain"

# Пароль
$password = $null
$cred = $null
if (-not $DryRun) {
    $cfgPassword = $config['credentials']['password']
    if ($cfgPassword -and $cfgPassword -ne 'PUT_PASSWORD_HERE') {
        $password = $cfgPassword
        Write-Log "Пароль: из конфига"
    } else {
        $c = Get-Credential -UserName $username -Message "SSH password for $username"
        if (-not $c) { Write-Log "Отменено пользователем" "ERROR"; exit 1 }
        $password = $c.GetNetworkCredential().Password
    }
    $secPwd = ConvertTo-SecureString $password -AsPlainText -Force
    $cred = New-Object System.Management.Automation.PSCredential($username, $secPwd)
}

# Inventory
Write-Log "Чтение inventory: $inventoryPath"
$inventory = Parse-SimpleYaml -Path $inventoryPath
$allHosts = Get-WindowsHosts -Inventory $inventory
$filteredHosts = Apply-HostsFilter -AllHosts $allHosts -HostFilter $HostFilter -HostsFile $HostsFile
Write-Log "Всего Windows-хостов: $($allHosts.Count)"
if ($HostsFile) { Write-Log "После HostsFile '$HostsFile': $($filteredHosts.Count)" } else { Write-Log "После фильтра '$HostFilter': $($filteredHosts.Count)" }
if ($filteredHosts.Count -eq 0) { Write-Log "Нет хостов для обработки" "WARN"; exit 0 }

$stateTransport = if ($Remove) { 'ssh-key-remove' } else { 'ssh-key' }
$stateDir = Get-StateDir -RootDir $ScriptDir -Domain $domain -Transport $stateTransport
Write-Log "State directory: $stateDir"
if ($Force) { Write-Log "Force: state будет игнорироваться" "WARN" }
if ($Remove) { Write-Log "Режим: REMOVE (ключ будет удалён из admin+user authorized_keys)" "WARN" }

if (-not $DryRun) {
    if (Test-WSManClientReady) {
        Write-Log "Локальный WSMan: TrustedHosts=*, Basic+AllowUnencrypted включены"
    } elseif (Test-IsAdmin -and (Initialize-LocalWSMan)) {
        Write-Log "Локальный WSMan донастроен (admin)"
    } else {
        Write-Log "Локальный WSMan не настроен — WinRM-fallback может не работать. Запусти .\Setup-LocalEnvironment.ps1 от админа." "WARN"
    }
}

Write-Log ""
Write-Log "=== Обработка ==="
Write-Log ""

$stats = @{
    ADDED = 0; SKIP_EXISTS = 0; SKIPPED = 0
    REMOVED = 0; NOT_FOUND = 0
    VERIFIED = 0; UNVERIFIED = 0
    SSH_OK_NO_AUTH = 0; SSH_AUTH_FAIL = 0
    WINRM_FAIL = 0; ERROR = 0
}

foreach ($h in $filteredHosts) {
    $displayName = "$($h.Name) ($($h.IP))"
    $statePath = Get-HostStatePath -StateDir $stateDir -HostName $h.Name
    $prevState = Read-HostState -Path $statePath

    if (-not $Remove -and (Test-ShouldSkip -State $prevState -Force:$Force)) {
        Write-Log "  [ SKIP ] $displayName — уже OK ($($prevState.timestamp))"
        $stats.SKIPPED++
        continue
    }

    $state = @{
        host = $h.Name; ip = $h.IP; transport = ''
        status = 'UNKNOWN'; category = ''; message = ''
        file = ''; is_admin = $false; verified = $null
    }

    $sshOpen = Test-TcpPort -ComputerName $h.IP -Port $sshPort -TimeoutSec $sshTimeout
    $winrmOpen = Test-TcpPort -ComputerName $h.IP -Port 5985 -TimeoutSec 2

    if ($DryRun) {
        $mode = if ($Remove) { 'REMOVE' } else { 'ADD' }
        $tag = if ($sshOpen) { 'SSH' } elseif ($winrmOpen) { 'WinRM (avans)' } else { 'NO_TRANSPORT' }
        Write-Log "  [ DRY:$mode ] $displayName — пойдёт через $tag (SSH:$sshOpen / WinRM:$winrmOpen)"
        continue
    }

    $r = $null
    # 1) SSH-парольный путь
    if ($sshOpen) {
        Write-Log "  [ssh] $displayName — пробуем через SSH"
        $r = Invoke-SshKeyDeploy -ComputerName $h.IP -Port $sshPort -Credential $cred `
            -PublicKey $publicKey -TargetUser $targetUser -TimeoutSec $sshTimeout `
            -RemoveMode $Remove.IsPresent
        $state.transport = 'ssh'

        if ($r.status -eq 'SSH_AUTH_FAIL' -and $winrmOpen) {
            Write-Log "  [ssh] парольная авторизация не сработала — fallback на WinRM" "WARN"
            $r = $null  # сбрасываем, попробуем WinRM
        }
    }

    # 2) WinRM-fallback (SSH порт закрыт ИЛИ SSH auth failed)
    if (-not $r) {
        if ($winrmOpen) {
            Write-Log "  [winrm] $displayName — пробуем через WinRM"
            $r = Invoke-WinRmKeyDeploy -ComputerName $h.IP -Credential $cred `
                -PublicKey $publicKey -TargetUser $targetUser `
                -RemoveMode $Remove.IsPresent
            $state.transport = 'winrm'
        } else {
            Write-Log "  [ FAIL ] $displayName — ни SSH, ни WinRM не доступны" "ERROR"
            $state.status = 'FAIL'; $state.category = 'NO_TRANSPORT'
            $state.message = 'Neither SSH (22) nor WinRM (5985) reachable'
            $state.transport = 'none'
            Write-HostState -Path $statePath -State $state
            $stats.ERROR++
            continue
        }
    }

    $state.status = $r.status
    $state.message = $r.message
    if ($r.file) { $state.file = $r.file }
    if ($null -ne $r.is_admin) { $state.is_admin = $r.is_admin }

    switch ($r.status) {
        'ADDED' {
            $role = if ($r.is_admin) { 'admin' } else { 'user' }
            Write-Log "  [ ADDED ] $displayName — $role -> $($r.file) (transport=$($state.transport))" "OK"
            $state.category = 'ADDED'
            $stats.ADDED++
        }
        'SKIP_EXISTS' {
            $role = if ($r.is_admin) { 'admin' } else { 'user' }
            Write-Log "  [ SKIP_EXISTS ] $displayName — ключ уже есть (${role}: $($r.file))"
            $state.category = 'SKIP_EXISTS'
            $stats.SKIP_EXISTS++
        }
        'REMOVED' {
            Write-Log "  [ REMOVED ] $displayName — удалён из: $($r.file) (transport=$($state.transport))" "OK"
            $state.category = 'REMOVED'
            $stats.REMOVED++
        }
        'NOT_FOUND' {
            Write-Log "  [ NOT_FOUND ] $displayName — ключ не найден ни в admin, ни в user authorized_keys"
            $state.category = 'NOT_FOUND'
            $stats.NOT_FOUND++
        }
        'ERROR' {
            Write-Log "  [ ERROR ] $displayName — $($r.message)" "ERROR"
            $state.category = 'REMOTE_ERROR'
            $stats.ERROR++
        }
        'SSH_AUTH_FAIL' {
            Write-Log "  [ SSH_AUTH_FAIL ] $displayName — $($r.message)" "ERROR"
            $state.category = 'SSH_AUTH_FAIL'
            $stats.SSH_AUTH_FAIL++
        }
        'WINRM_CONNECT_FAIL' {
            Write-Log "  [ WINRM_FAIL ] $displayName — $($r.message)" "ERROR"
            $state.category = 'WINRM_CONNECT_FAIL'
            $stats.WINRM_FAIL++
        }
        'WINRM_INVOKE_FAIL' {
            Write-Log "  [ WINRM_FAIL ] $displayName — $($r.message)" "ERROR"
            $state.category = 'WINRM_INVOKE_FAIL'
            $stats.WINRM_FAIL++
        }
        default {
            Write-Log "  [ ? ] $displayName — статус $($r.status): $($r.message)" "WARN"
            $state.category = $r.status
            $stats.ERROR++
        }
    }
    if ($r.acl_err) { Write-Log "    ACL warning: $($r.acl_err)" "WARN" }

    # Post-check — только в Add-режиме, если ключ положен и SSH порт открыт
    if (-not $Remove -and ($r.status -eq 'ADDED' -or $r.status -eq 'SKIP_EXISTS') -and $postCheckEnabled -and $sshOpen) {
        $v = Test-SshKeyLogin -ComputerName $h.IP -Port $sshPort -Username $username `
            -PrivateKeyPath $privateKeyPath -TimeoutSec $sshTimeout
        if ($v.Ok) {
            Write-Log "    [ verify ] вход по ключу работает" "OK"
            $state.verified = $true
            $stats.VERIFIED++
        } else {
            Write-Log "    [ verify ] вход по ключу НЕ работает: $($v.Error)" "WARN"
            $state.verified = $false
            $stats.UNVERIFIED++
        }
    } elseif (($r.status -eq 'ADDED' -or $r.status -eq 'SKIP_EXISTS') -and -not $sshOpen) {
        Write-Log "    [ verify ] post-check пропущен (SSH порт закрыт — ключ положен авансом через WinRM)" "WARN"
        $state.verified = $false
        $stats.SSH_OK_NO_AUTH++
    }

    Write-HostState -Path $statePath -State $state
}

# --- Итоги ---
Write-Log ""
Write-Log "=========================================="
Write-Log "=== ИТОГО ==="
Write-Log "=========================================="
Write-Log "Всего хостов:           $($filteredHosts.Count)"
if ($Remove) {
    Write-Log "Ключ удалён:            $($stats.REMOVED)"
    Write-Log "Не было ключа:          $($stats.NOT_FOUND)"
} else {
    Write-Log "Ключ добавлен:          $($stats.ADDED)"
    Write-Log "Уже был (skip-exists):  $($stats.SKIP_EXISTS)"
    Write-Log "Skipped (state OK):     $($stats.SKIPPED)"
    if ($postCheckEnabled) {
        Write-Log "Вход по ключу OK:       $($stats.VERIFIED)"
        Write-Log "Вход по ключу FAIL:     $($stats.UNVERIFIED)"
        Write-Log "Ключ авансом (no SSH):  $($stats.SSH_OK_NO_AUTH)"
    }
}
Write-Log "SSH auth fail:          $($stats.SSH_AUTH_FAIL)"
Write-Log "WinRM failed:           $($stats.WINRM_FAIL)"
Write-Log "Ошибки:                 $($stats.ERROR)"
Write-Log ""
Write-Log "State: $stateDir"
Write-Log "Лог: $LogFile"

if ($stats.ERROR -gt 0 -or $stats.WINRM_FAIL -gt 0 -or $stats.SSH_AUTH_FAIL -gt 0) { exit 1 }

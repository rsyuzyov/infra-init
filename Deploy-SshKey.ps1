<#
.SYNOPSIS
    Массовое прописывание SSH публичного ключа на Windows-хостах через SSH.
.DESCRIPTION
    Запускать после Deploy-OpenSSH, когда на хостах уже работает OpenSSH-сервер
    с парольной авторизацией. Скрипт подключается по SSH (Posh-SSH), на хосте
    через PowerShell:
      - определяет, является ли target_user членом BUILTIN\Administrators
      - кладёт публичный ключ в правильный файл:
          админ → C:\ProgramData\ssh\administrators_authorized_keys
          юзер  → C:\Users\<short_name>\.ssh\authorized_keys
      - идемпотентно (не дублирует ключ — сравнение по телу)
      - выставляет ACL согласно требованиям sshd
.PARAMETER DryRun
    Только показать список хостов и доступность SSH, без изменений.
.PARAMETER HostFilter
    Фильтр по имени хоста или IP (wildcard *).
.PARAMETER ConfigPath
    Путь к config.yaml. По умолчанию — рядом со скриптом.
.PARAMETER PublicKeyPath
    Переопределяет public_key_path из config.yaml.
.EXAMPLE
    .\Deploy-SshKey.ps1 -DryRun
    .\Deploy-SshKey.ps1 -HostFilter "DC*"
    .\Deploy-SshKey.ps1
.NOTES
    Требует модуль Posh-SSH:
      Install-Module -Name Posh-SSH -Scope CurrentUser -Force
#>

param(
    [switch]$DryRun,
    [string]$HostFilter = "*",
    [string]$ConfigPath,
    [string]$PublicKeyPath
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$LogFile = Join-Path $ScriptDir "Deploy-SshKey.log"

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
            if ($val -eq '{}') { $current[$key] = @{} } else { $current[$key] = $val }
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
    param([hashtable]$Node, [string]$Name)
    if (-not $Node) { return $null }
    if ($Node.ContainsKey($Name) -and $Node[$Name] -is [hashtable]) { return $Node[$Name] }
    foreach ($key in $Node.Keys) {
        $val = $Node[$key]
        if ($val -is [hashtable] -and $val['children'] -is [hashtable]) {
            $found = Find-GroupByName -Node $val['children'] -Name $Name
            if ($found) { return $found }
        }
    }
    return $null
}

function Collect-HostsRecursive {
    param([hashtable]$Group, [System.Collections.ArrayList]$Acc)
    if (-not $Group) { return }
    if ($Group['hosts'] -is [hashtable]) {
        foreach ($hostName in $Group['hosts'].Keys) {
            $hostData = $Group['hosts'][$hostName]
            if ($hostData -is [hashtable] -and $hostData['ansible_host']) {
                [void]$Acc.Add(@{ Name = $hostName; IP = $hostData['ansible_host'] })
            }
        }
    }
    if ($Group['children'] -is [hashtable]) {
        foreach ($childName in $Group['children'].Keys) {
            $child = $Group['children'][$childName]
            if ($child -is [hashtable]) { Collect-HostsRecursive -Group $child -Acc $Acc }
        }
    }
}

function Get-WindowsHosts {
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

function Test-SshPort {
    param([string]$ComputerName, [int]$Port = 22, [int]$TimeoutSec = 5)
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $connect = $tcp.BeginConnect($ComputerName, $Port, $null, $null)
        $wait = $connect.AsyncWaitHandle.WaitOne($TimeoutSec * 1000, $false)
        if (-not $wait) { $tcp.Close(); return $false }
        $tcp.EndConnect($connect)
        $tcp.Close()
        return $true
    } catch { return $false }
}

function Test-SshKeyLogin {
    # Пробует подключиться по приватному ключу. Возвращает hashtable @{ Ok=$true|$false; Error='...' }.
    # Posh-SSH требует -Credential (PSCredential); для key-only передаём dummy пароль —
    # если ключ примут, пароль игнорируется; если не примут, пароль тоже не сработает и мы получим fail.
    param(
        [string]$ComputerName,
        [int]$Port,
        [string]$Username,
        [string]$PrivateKeyPath,
        [int]$TimeoutSec
    )
    $dummyPwd = ConvertTo-SecureString "no-such-password-key-only" -AsPlainText -Force
    $keyCred = New-Object System.Management.Automation.PSCredential($Username, $dummyPwd)
    $session = $null
    try {
        $session = New-SSHSession -ComputerName $ComputerName -Port $Port -Credential $keyCred `
            -KeyFile $PrivateKeyPath -AcceptKey -ConnectionTimeout $TimeoutSec -ErrorAction Stop
        return @{ Ok = $true; Error = "" }
    } catch {
        return @{ Ok = $false; Error = $_.Exception.Message }
    } finally {
        if ($session) { Remove-SSHSession -SessionId $session.SessionId -ErrorAction SilentlyContinue | Out-Null }
    }
}

# ============================================================
# Удалённый скрипт. Литеральная here-string + .Replace() —
# без эскейпов и без интерполяции на стороне сборки.
# Запускается через `powershell -EncodedCommand <base64>`.
# Stdout: одна строка между маркерами BEGIN/END с JSON-результатом.
# ============================================================

function Build-RemoteScript {
    param([string]$PublicKey, [string]$TargetUser)

    $pkEsc = $PublicKey -replace "'", "''"
    $tuEsc = $TargetUser -replace "'", "''"

    $template = @'
$ErrorActionPreference = 'Stop'
$result = @{ Status = 'ERROR'; Message = ''; File = ''; IsAdmin = $false; AclErr = $null }
try {
    $publicKey  = '__PUBKEY__'
    $targetUser = '__TARGETUSER__'

    $acct = New-Object System.Security.Principal.NTAccount($targetUser)
    $userSid = $acct.Translate([System.Security.Principal.SecurityIdentifier]).Value

    $isAdmin = $false
    $members = Get-LocalGroupMember -SID 'S-1-5-32-544' -ErrorAction Stop
    foreach ($m in $members) {
        if ($m.SID.Value -eq $userSid) { $isAdmin = $true; break }
    }

    $shortName = $targetUser
    if ($shortName -match '\\') { $shortName = $shortName.Split('\')[-1] }

    if ($isAdmin) {
        $sshDir = 'C:\ProgramData\ssh'
        $keyFile = Join-Path $sshDir 'administrators_authorized_keys'
    } else {
        $profilePath = "C:\Users\$shortName"
        if (-not (Test-Path $profilePath)) {
            throw "Profile not found: $profilePath (non-admin user requires existing profile)"
        }
        $sshDir = Join-Path $profilePath '.ssh'
        $keyFile = Join-Path $sshDir 'authorized_keys'
    }
    $result.File = $keyFile
    $result.IsAdmin = $isAdmin

    if (-not (Test-Path $sshDir)) {
        New-Item -ItemType Directory -Path $sshDir -Force | Out-Null
    }

    $keyParts = $publicKey.Trim() -split '\s+'
    if ($keyParts.Count -lt 2) { throw 'Invalid public key format' }
    $keyBody = $keyParts[1]

    $alreadyExists = $false
    if (Test-Path $keyFile) {
        $lines = Get-Content $keyFile -ErrorAction SilentlyContinue
        foreach ($l in $lines) {
            $l = $l.Trim()
            if (-not $l -or $l.StartsWith('#')) { continue }
            $parts = $l -split '\s+'
            if ($parts.Count -ge 2 -and $parts[1] -eq $keyBody) {
                $alreadyExists = $true
                break
            }
        }
    }

    if ($alreadyExists) {
        $result.Status = 'SKIP_EXISTS'
    } else {
        Add-Content -Path $keyFile -Value $publicKey.Trim() -Encoding ASCII
        $result.Status = 'ADDED'
    }

    try {
        if ($isAdmin) {
            $null = & icacls $keyFile /inheritance:r 2>&1
            $null = & icacls $keyFile /grant 'BUILTIN\Administrators:F' 'NT AUTHORITY\SYSTEM:F' 2>&1
        } else {
            $null = & icacls $keyFile /inheritance:r 2>&1
            $grantArg = $shortName + ':F'
            $null = & icacls $keyFile /grant $grantArg 'NT AUTHORITY\SYSTEM:F' 2>&1
        }
    } catch {
        $result.AclErr = $_.Exception.Message
    }
} catch {
    $result.Status = 'ERROR'
    $result.Message = $_.Exception.Message
}
'__SSHKEYDEPLOY_BEGIN__' + ($result | ConvertTo-Json -Compress) + '__SSHKEYDEPLOY_END__'
'@

    return $template.Replace('__PUBKEY__', $pkEsc).Replace('__TARGETUSER__', $tuEsc)
}

function Encode-PsCommand {
    param([string]$Script)
    $bytes = [System.Text.Encoding]::Unicode.GetBytes($Script)
    return [Convert]::ToBase64String($bytes)
}

function Parse-RemoteResult {
    param([string]$Output)
    $m = [regex]::Match($Output, '__SSHKEYDEPLOY_BEGIN__(?<json>.*?)__SSHKEYDEPLOY_END__', 'Singleline')
    if (-not $m.Success) { return $null }
    try { return $m.Groups['json'].Value | ConvertFrom-Json } catch { return $null }
}

# ============================================================
# MAIN
# ============================================================

Write-Log "=========================================="
Write-Log "=== Deploy-SshKey - Start ==="
Write-Log "=========================================="

if (-not (Get-Module -ListAvailable -Name Posh-SSH)) {
    Write-Log "Module Posh-SSH is not installed." "ERROR"
    Write-Log "Install: Install-Module -Name Posh-SSH -Scope CurrentUser -Force" "ERROR"
    exit 1
}
Import-Module Posh-SSH -ErrorAction Stop | Out-Null

if (-not $ConfigPath) { $ConfigPath = Join-Path $ScriptDir "config.yaml" }
if (-not (Test-Path $ConfigPath)) {
    Write-Log "Config not found: $ConfigPath" "ERROR"
    exit 1
}
Write-Log "Config: $ConfigPath"
$config = Parse-SimpleYaml -Path $ConfigPath

$inventoryPath = $config['inventory_path']
$username      = $config['credentials']['username']
$targetUser    = $config['target_user']
$pubKeyPath    = if ($PublicKeyPath) { $PublicKeyPath } else { $config['public_key_path'] }
$sshPort       = if ($config['ssh_port']) { [int]$config['ssh_port'] } else { 22 }
$sshTimeout    = if ($config['ssh_timeout']) { [int]$config['ssh_timeout'] } else { 10 }

if (-not $targetUser) { $targetUser = $username }

if (-not $pubKeyPath -or -not (Test-Path $pubKeyPath)) {
    Write-Log "Public key not found: $pubKeyPath" "ERROR"
    exit 1
}
if (-not (Test-Path $inventoryPath)) {
    Write-Log "Inventory not found: $inventoryPath" "ERROR"
    exit 1
}

# Приватный ключ для post-check. Если не задан — выводим из public_key_path,
# отрезая .pub. Если приватный ключ не найден — post-check пропускается.
$privateKeyPath = $config['private_key_path']
if (-not $privateKeyPath) {
    if ($pubKeyPath.ToLower().EndsWith('.pub')) {
        $privateKeyPath = $pubKeyPath.Substring(0, $pubKeyPath.Length - 4)
    }
}
$postCheckEnabled = $false
if ($privateKeyPath) {
    if (Test-Path $privateKeyPath) {
        $postCheckEnabled = $true
    } else {
        Write-Log "Private key not found: $privateKeyPath - post-check will be skipped" "WARN"
    }
}

$publicKey = (Get-Content $pubKeyPath -Raw).Trim()
$keyParts = $publicKey -split '\s+'
if (-not $publicKey -or $keyParts.Count -lt 2) {
    Write-Log "Key file is empty or not authorized_keys format: $pubKeyPath" "ERROR"
    exit 1
}
$blob = $keyParts[1]
$fp = $blob.Substring(0, [Math]::Min(16, $blob.Length)) + "..."
Write-Log "Public key: $pubKeyPath ($fp)"
if ($postCheckEnabled) { Write-Log "Private key for post-check: $privateKeyPath" }
Write-Log "SSH login: $username, target user on hosts: $targetUser"
Write-Log "SSH port: $sshPort, timeout: ${sshTimeout}s"

if (-not $DryRun) {
    $cfgPassword = $config['credentials']['password']
    if ($cfgPassword -and $cfgPassword -ne 'PUT_PASSWORD_HERE') {
        Write-Log "Password: from config"
        $password = $cfgPassword
    } else {
        Write-Log "Asking for credentials of $username..."
        $cred = Get-Credential -UserName $username -Message "SSH password for $username"
        if (-not $cred) { Write-Log "Cancelled by user" "ERROR"; exit 1 }
        $password = $cred.GetNetworkCredential().Password
    }
    $secPwd = ConvertTo-SecureString $password -AsPlainText -Force
    $sshCred = New-Object System.Management.Automation.PSCredential($username, $secPwd)
}

Write-Log "Reading inventory: $inventoryPath"
$inventory = Parse-SimpleYaml -Path $inventoryPath
$allHosts = Get-WindowsHosts -Inventory $inventory

$filteredHosts = $allHosts | Where-Object { $_.Name -like $HostFilter -or $_.IP -like $HostFilter }
$filteredHosts = $filteredHosts | Sort-Object { [version]($_.IP -replace '(\d+)\.(\d+)\.(\d+)\.(\d+)', '$1.$2.$3.$4') }

Write-Log "Total Windows hosts: $($allHosts.Count)"
Write-Log "After filter '$HostFilter': $($filteredHosts.Count)"

if ($filteredHosts.Count -eq 0) {
    Write-Log "No hosts to process" "WARN"
    exit 0
}

$remoteScript = Build-RemoteScript -PublicKey $publicKey -TargetUser $targetUser
$encoded = Encode-PsCommand -Script $remoteScript
$remoteCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -EncodedCommand $encoded"

Write-Log ""
Write-Log "=== Deploying key ==="
Write-Log ""

$stats = @{ ADDED = 0; SKIP_EXISTS = 0; SSH_FAIL = 0; ERROR = 0; VERIFIED = 0; UNVERIFIED = 0 }
$problems = @()

foreach ($h in $filteredHosts) {
    $displayName = "$($h.Name) ($($h.IP))"

    if (-not (Test-SshPort -ComputerName $h.IP -Port $sshPort -TimeoutSec $sshTimeout)) {
        Write-Log "  [ SSH_FAIL ] $displayName - port $sshPort unreachable" "WARN"
        $stats.SSH_FAIL++
        $problems += @{ Host = $displayName; Reason = "port $sshPort unreachable" }
        continue
    }

    if ($DryRun) {
        Write-Log "  [ DRY ] $displayName - would add key for $targetUser"
        continue
    }

    $session = $null
    try {
        $session = New-SSHSession -ComputerName $h.IP -Port $sshPort -Credential $sshCred -AcceptKey -ConnectionTimeout $sshTimeout -ErrorAction Stop
        $cmdResult = Invoke-SSHCommand -SessionId $session.SessionId -Command $remoteCmd -TimeOut 60 -ErrorAction Stop

        $parsed = Parse-RemoteResult -Output ($cmdResult.Output -join "`n")
        if (-not $parsed) {
            $tail = ($cmdResult.Output -join ' | ')
            if ($tail.Length -gt 300) { $tail = $tail.Substring(0, 300) + '...' }
            Write-Log "  [ ERROR ] $displayName - cannot parse output: $tail" "ERROR"
            $stats.ERROR++
            $problems += @{ Host = $displayName; Reason = "parse_failed" }
            continue
        }

        switch ($parsed.Status) {
            "ADDED" {
                $role = if ($parsed.IsAdmin) { "admin" } else { "user" }
                Write-Log "  [ ADDED ] $displayName - ${role} -> $($parsed.File)" "OK"
                $stats.ADDED++
            }
            "SKIP_EXISTS" {
                $role = if ($parsed.IsAdmin) { "admin" } else { "user" }
                Write-Log "  [ SKIP  ] $displayName - key already present (${role}: $($parsed.File))"
                $stats.SKIP_EXISTS++
            }
            "ERROR" {
                Write-Log "  [ ERROR ] $displayName - $($parsed.Message)" "ERROR"
                $stats.ERROR++
                $problems += @{ Host = $displayName; Reason = $parsed.Message }
            }
            default {
                Write-Log "  [ ERROR ] $displayName - unknown status: $($parsed.Status)" "ERROR"
                $stats.ERROR++
                $problems += @{ Host = $displayName; Reason = "unknown_status:$($parsed.Status)" }
            }
        }
        if ($parsed.AclErr) { Write-Log "    ACL warning: $($parsed.AclErr)" "WARN" }

        # Post-check: ключ записан (или уже был) - проверим, что заход реально работает.
        if ($postCheckEnabled -and ($parsed.Status -eq 'ADDED' -or $parsed.Status -eq 'SKIP_EXISTS')) {
            $verify = Test-SshKeyLogin -ComputerName $h.IP -Port $sshPort -Username $username `
                -PrivateKeyPath $privateKeyPath -TimeoutSec $sshTimeout
            if ($verify.Ok) {
                Write-Log "    [ verify ] key login works" "OK"
                $stats.VERIFIED++
            } else {
                Write-Log "    [ verify ] key login FAILED: $($verify.Error)" "WARN"
                $stats.UNVERIFIED++
                $problems += @{ Host = $displayName; Reason = "key login failed: $($verify.Error)" }
            }
        }
    } catch {
        Write-Log "  [ ERROR ] $displayName - $($_.Exception.Message)" "ERROR"
        $stats.ERROR++
        $problems += @{ Host = $displayName; Reason = $_.Exception.Message }
    } finally {
        if ($session) { Remove-SSHSession -SessionId $session.SessionId -ErrorAction SilentlyContinue | Out-Null }
    }
}

Write-Log ""
Write-Log "=========================================="
Write-Log "=== SUMMARY ==="
Write-Log "=========================================="
Write-Log "Total hosts:       $($filteredHosts.Count)"
Write-Log "Added:             $($stats.ADDED)"
Write-Log "Already present:   $($stats.SKIP_EXISTS)"
Write-Log "SSH unreachable:   $($stats.SSH_FAIL)"
Write-Log "Errors:            $($stats.ERROR)"
if ($postCheckEnabled) {
    Write-Log "Key login OK:      $($stats.VERIFIED)"
    Write-Log "Key login FAILED:  $($stats.UNVERIFIED)"
}

if ($problems.Count -gt 0) {
    Write-Log ""
    Write-Log "Problem hosts:"
    foreach ($p in $problems) { Write-Log "  - $($p.Host): $($p.Reason)" "WARN" }
}

if ($stats.ERROR -gt 0) { exit 1 }

Write-Log ""
Write-Log "=== Done ===" "OK"
Write-Log "Log: $LogFile"

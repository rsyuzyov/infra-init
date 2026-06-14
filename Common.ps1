<#
.SYNOPSIS
    Общие функции для Deploy-*.ps1 скриптов в infra-init.
.DESCRIPTION
    Импортируется через dot-sourcing:
        . (Join-Path $PSScriptRoot "Common.ps1")

    Содержит:
      - Write-Log              — логирование в файл $LogFile (из caller scope) + Host
      - New-LogFile            — создание log/<script>/<timestamp>.log
      - Parse-SimpleYaml       — мини-парсер YAML
      - Find-GroupByName       — рекурсивный поиск группы в inventory
      - Collect-HostsRecursive — рекурсивный сбор хостов
      - Get-WindowsHosts       — все Windows-хосты из inventory
      - Read-HostsFile         — список хостов из текстового файла
      - Apply-HostsFilter      — фильтрация $allHosts по HostsFile или HostFilter
      - Test-TcpPort           — проверка TCP-порта
      - State machine:
          Get-StateDir, Get-HostStatePath, Read-HostState, Write-HostState, Test-ShouldSkip
#>

# ============================================================
# Admin / elevation
# ============================================================

function Test-IsAdmin {
    $current = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($current)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Invoke-SelfElevation {
    # Перезапускает скрипт от админа с теми же параметрами и сохраняет current working directory.
    # ScriptPath — полный путь .ps1; BoundParams — $PSBoundParameters caller'а.
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][System.Collections.IDictionary]$BoundParams
    )
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-NoExit', '-File', "`"$ScriptPath`"")
    foreach ($k in $BoundParams.Keys) {
        $v = $BoundParams[$k]
        if ($v -is [System.Management.Automation.SwitchParameter]) {
            if ($v.IsPresent) { $argList += "-$k" }
        } elseif ($v -is [bool]) {
            if ($v) { $argList += "-$k" }
        } else {
            $argList += "-$k"
            $argList += "`"$v`""
        }
    }
    $wd = (Get-Location).Path
    try {
        Start-Process powershell.exe -Verb RunAs -WorkingDirectory $wd -ArgumentList $argList -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function Assert-Admin {
    # Проверяет права админа. Если их нет — пробует поднять окно через UAC.
    # Если пользователь отменил UAC — печатает инструкцию и завершает скрипт.
    # Если задан -SkipIf и он вернул $true — не требуем админа (например, всё уже настроено).
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][System.Collections.IDictionary]$BoundParams,
        [scriptblock]$SkipIf,
        [string]$Reason = "управление WinRM/PsExec"
    )
    if (Test-IsAdmin) { return }
    if ($SkipIf) {
        try { if (& $SkipIf) { return } } catch { }
    }

    Write-Host ""
    Write-Host "Этот скрипт требует прав администратора ($Reason)." -ForegroundColor Yellow
    Write-Host "Запрашиваю повышение через UAC..." -ForegroundColor Yellow

    if (Invoke-SelfElevation -ScriptPath $ScriptPath -BoundParams $BoundParams) {
        Write-Host "Скрипт перезапущен в админ-сессии. Текущее окно можно закрыть." -ForegroundColor Green
        exit 0
    }

    Write-Host ""
    Write-Host "Не удалось поднять права (UAC отменён или заблокирован)." -ForegroundColor Red
    Write-Host "Запусти PowerShell от имени администратора и снова:" -ForegroundColor Red
    Write-Host "    $ScriptPath" -ForegroundColor Red
    Write-Host ""
    exit 1
}

function Test-WSManClientReady {
    # $true, если локальный WSMan-клиент уже настроен для Invoke-Command к рабочим группам:
    # служба запущена, TrustedHosts=*, Auth\Basic=true, AllowUnencrypted=true.
    try {
        $svc = Get-Service WinRM -ErrorAction Stop
        if ($svc.Status -ne 'Running') { return $false }
        $th = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction Stop).Value
        if ($th -ne '*') { return $false }
        $basic = (Get-Item WSMan:\localhost\Client\Auth\Basic -ErrorAction Stop).Value
        if (-not $basic) { return $false }
        $unenc = (Get-Item WSMan:\localhost\Client\AllowUnencrypted -ErrorAction Stop).Value
        if (-not $unenc) { return $false }
        return $true
    } catch {
        return $false
    }
}

function Initialize-LocalWSMan {
    # Конфигурирует локальный WSMan-клиент для Invoke-Command к хостам в рабочих группах.
    # Требует прав админа. При неудаче — возвращает $false (caller сам решит, что делать).
    try {
        $svc = Get-Service WinRM -ErrorAction SilentlyContinue
        if (-not $svc) { return $false }
        if ($svc.Status -ne 'Running') { Start-Service WinRM -ErrorAction Stop }
        Set-Item WSMan:\localhost\Client\AllowUnencrypted -Value $true -Force -ErrorAction SilentlyContinue
        Set-Item WSMan:\localhost\Client\Auth\Basic -Value $true -Force -ErrorAction SilentlyContinue
        $current = (Get-Item WSMan:\localhost\Client\TrustedHosts).Value
        if ($current -ne '*') {
            Set-Item WSMan:\localhost\Client\TrustedHosts -Value '*' -Force -ErrorAction SilentlyContinue
        }
        return $true
    } catch {
        return $false
    }
}

# ============================================================
# Логирование
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
    if ($LogFile) {
        Add-Content -Path $LogFile -Value $entry -Encoding UTF8 -ErrorAction SilentlyContinue
    }
}

function New-LogFile {
    # Создаёт log/<script>/<timestamp>.log относительно $RootDir и возвращает абсолютный путь.
    param(
        [string]$RootDir,
        [string]$ScriptName
    )
    $stamp = (Get-Date -Format "yyyy-MM-ddTHHmmss")
    $dir = Join-Path $RootDir (Join-Path "log" $ScriptName)
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    return (Join-Path $dir "$stamp.log")
}

# ============================================================
# YAML parser + inventory
# ============================================================

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
            $val = $Matches[2]
            # Срезаем inline-комментарий: '#' после whitespace
            $commentIdx = -1
            for ($i = 0; $i -lt $val.Length; $i++) {
                if ($val[$i] -eq '#' -and ($i -eq 0 -or $val[$i-1] -eq ' ' -or $val[$i-1] -eq "`t")) {
                    $commentIdx = $i; break
                }
            }
            if ($commentIdx -ge 0) { $val = $val.Substring(0, $commentIdx) }
            $val = $val.Trim().Trim('"').Trim("'")
            if (-not $val) {
                # Значение оказалось пустым (был только комментарий) — трактуем как пустой узел
                $newObj = @{}
                $current[$key] = $newObj
                $stack += @{ obj = $newObj; indent = $indent }
            }
            elseif ($val -eq '{}') { $current[$key] = @{} }
            else { $current[$key] = $val }
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

# ============================================================
# Hosts filter
# ============================================================

function Read-HostsFile {
    # Читает файл с хостами (по одному имени/IP на строку, # — комментарий).
    # Возвращает hashtable @{ <lowercased-line> = $true } для быстрого matching.
    param([string]$Path)
    $wanted = @{}
    foreach ($line in Get-Content $Path) {
        $line = $line.Trim()
        if (-not $line -or $line.StartsWith('#')) { continue }
        $wanted[$line.ToLower()] = $true
    }
    return $wanted
}

function Apply-HostsFilter {
    # Фильтрует $allHosts: HostsFile приоритетнее HostFilter.
    # Сортирует результат по IP.
    param(
        [array]$AllHosts,
        [string]$HostFilter = '*',
        [string]$HostsFile
    )
    if ($HostsFile) {
        $wanted = Read-HostsFile -Path $HostsFile
        $filtered = $AllHosts | Where-Object {
            $wanted.ContainsKey($_.Name.ToLower()) -or $wanted.ContainsKey($_.IP.ToLower())
        }
    } else {
        $filtered = $AllHosts | Where-Object {
            $_.Name -like $HostFilter -or $_.IP -like $HostFilter
        }
    }
    return @($filtered | Sort-Object {
        [version]($_.IP -replace '(\d+)\.(\d+)\.(\d+)\.(\d+)', '$1.$2.$3.$4')
    })
}

# ============================================================
# Network probes
# ============================================================

function Test-TcpPort {
    param([string]$ComputerName, [int]$Port, [int]$TimeoutSec = 3)
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

# ============================================================
# Remote access-denied (UAC remote token filtering) helpers
# ============================================================

function Test-AccessDenied {
    # $true, если вывод транспорта означает «отказано в доступе».
    # Подходит и для PsExec (stderr), и для WinRM (текст исключения).
    # Частая причина под локальной учёткой-админом — UAC remote token filtering:
    # удалённый вход даёт урезанный токен без админ-прав.
    # ВАЖНО: ловим именно отказ доступа, а не общий префикс PsExec "Couldn't access <host>:" —
    # он же появляется при закрытом 445 ("сетевой путь не найден") и неверном пароле
    # ("logon failure"), где подсказка про LocalAccountTokenFilterPolicy неуместна.
    # Признаки именно access denied: код 5 (ERROR_ACCESS_DENIED) и текст ошибки (EN/RU).
    # RU-текст из PsExec читается корректно, т.к. stderr читается с -Encoding Oem (см. вызывающий код).
    param([string]$Message, [int]$ExitCode = 0)
    if ($ExitCode -eq 5) { return $true }
    if (-not $Message) { return $false }
    return ($Message -match 'Access is denied' -or
            $Message -match 'Отказано в доступе')
}

function Write-AccessDeniedHint {
    # Справочная подсказка при отказе доступа: как разрешить полный токен
    # локальным админам на целевом хосте (UAC remote token filtering).
    param([string]$ComputerName)
    Write-Log "  ─── Отказано в доступе — вероятно, UAC remote token filtering ───" "WARN"
    Write-Log "  Логин/сессия проходят, но локальный админ при удалённом входе получает" "WARN"
    Write-Log '  урезанный токен: PsExec не достучится до admin$/SCM, а команды внутри' "WARN"
    Write-Log "  WinRM-сессии падают с «Отказано в доступе». Лечится на ЦЕЛЕВОМ хосте ($ComputerName):" "WARN"
    Write-Log '    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v LocalAccountTokenFilterPolicy /t REG_DWORD /d 1 /f' "WARN"
    Write-Log "  Применяется сразу, перезагрузка не нужна. Не нужно для доменных учёток" "WARN"
    Write-Log "  и встроенного Administrator." "WARN"
    Write-Log "  ─────────────────────────────────────────────────────────────────────" "WARN"
}

# ============================================================
# State machine (output/<domain>/<transport>/<host>.json)
# ============================================================

function Get-StateDir {
    # output/<domain>/<transport>/, создаёт каталог.
    param(
        [string]$RootDir,
        [string]$Domain,
        [string]$Transport
    )
    if (-not $Domain) { $Domain = '_default' }
    $dir = Join-Path $RootDir (Join-Path "output" (Join-Path $Domain $Transport))
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    return $dir
}

function Get-HostStatePath {
    # Имя файла берётся из имени хоста с заменой запрещённых символов.
    param([string]$StateDir, [string]$HostName)
    $safe = $HostName -replace '[\\/:*?"<>|]', '_'
    return (Join-Path $StateDir "$safe.json")
}

function Read-HostState {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    try {
        return Get-Content $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        return $null
    }
}

function Write-HostState {
    # Сохраняет результат обработки хоста.
    param(
        [string]$Path,
        [hashtable]$State   # Ожидаем поля: host, ip, transport, status, category, exit_code, message
    )
    $State['timestamp'] = (Get-Date -Format "yyyy-MM-ddTHH:mm:ss")
    $json = $State | ConvertTo-Json -Depth 5 -Compress:$false
    Set-Content -Path $Path -Value $json -Encoding UTF8
}

function Test-ShouldSkip {
    # true, если хост уже обработан успешно и -Force не задан.
    # Успехом считаем status = OK или SKIP_EXISTS.
    param(
        $State,         # объект из Read-HostState или $null
        [bool]$Force = $false
    )
    if ($Force) { return $false }
    if (-not $State) { return $false }
    if (-not $State.status) { return $false }
    return ($State.status -eq 'OK' -or $State.status -eq 'SKIP_EXISTS')
}

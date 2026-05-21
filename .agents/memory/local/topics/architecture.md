# Архитектура infra-init

## Структура каталогов

```
infra-init/
├── Common.ps1                  ← общий модуль (логи, YAML, inventory, TCP-probe, state machine, admin/WSMan)
├── Setup-LocalEnvironment.ps1  ← разовая настройка локального WSMan-клиента (под админом)
├── Deploy-WinRM.ps1            ← WinRM через PsExec
├── Deploy-OpenSSH.ps1          ← OpenSSH: WinRM-first, PsExec-fallback, MSI-fallback
├── Deploy-SshKey.ps1           ← SSH-ключ: SSH-first (Posh-SSH), WinRM-fallback (Invoke-Command), -Remove
├── config.yaml                 ← реальные значения (gitignored)
├── config.yaml.example         ← шаблон
├── tools/                      ← PsExec.exe, OpenSSH-Win64.msi
├── log/<script>/<ts>.log       ← общий лог запуска (gitignored)
├── output/<domain>/<transport>/<host>.json ← per-host state machine
└── tmp/                        ← временные файлы (gitignored)
```

## Общая библиотека `Common.ps1`

Импорт во все Deploy-*.ps1 и Setup-LocalEnvironment.ps1:
```powershell
. (Join-Path $PSScriptRoot "Common.ps1")
```

Внутри:
- `Write-Log` — пишет в `$LogFile` (из caller scope) + Host.
- `New-LogFile -RootDir -ScriptName` — создаёт `log/<script>/<timestamp>.log` и возвращает путь.
- `Parse-SimpleYaml` — мини-парсер.
- `Find-GroupByName`, `Collect-HostsRecursive`, `Get-WindowsHosts` — извлечение Windows-хостов из inventory.
- `Read-HostsFile`, `Apply-HostsFilter` — фильтр по `-HostsFile` или `-HostFilter`.
- `Test-TcpPort` — TCP-probe с таймаутом.
- State machine: `Get-StateDir`, `Get-HostStatePath`, `Read-HostState`, `Write-HostState`, `Test-ShouldSkip`.

## State machine

Per-host JSON в `output/<domain>/<transport>/<host>.json`. Структура:
```json
{
  "host": "<short-name>", "ip": "<ip>",
  "transport": "winrm" | "psexec" | "ssh" | "ssh-msi" | "winrm-msi" | "none",
  "status": "OK" | "SKIP_EXISTS" | "REBOOT_REQUIRED" | "SMB_FAIL" | "OFFLINE" | "FAIL" | ...,
  "category": "ALREADY_WORKING" | "CONFIGURED" | "PSEXEC_ERROR" | ...,
  "exit_code": 0,
  "message": "...",
  "timestamp": "2026-05-14T15:12:14"
}
```

`Test-ShouldSkip` пропускает хост, если `status in {OK, SKIP_EXISTS}` и не задан `-Force`.

## `domain` в config.yaml

Произвольный идентификатор окружения (`mycorp.local`, `staging`, `client-a`, ...). Используется для разделения `output/<domain>/` между разными контекстами/клиентами. По умолчанию `_default`.

⚠️ В соседнем проекте `net-conf-gen` уже введён аналогичный domain: его inventory теперь лежит в `output/<domain>/inventory.yaml`. В `infra-init/config.yaml` `inventory_path` должен указывать на новый путь — например `<repo-root>\net-conf-gen\output\<domain>\inventory.yaml`.

## Транспортная стратегия

### Deploy-WinRM
PsExec — единственный путь (потому что WinRM ещё не настроен). Pre-check TCP 445 → если закрыт, мгновенный `SMB_FAIL` без таймаута. Категории: `ALREADY_WORKING`, `CONFIGURED`, `SMB_UNREACHABLE`, `PSEXEC_ERROR`, `VERIFY_FAIL`, `OFFLINE`.

### Deploy-OpenSSH (WinRM-first)
1. **WinRM** (`Invoke-Command`, нативный scriptblock, hashtable-результат) — основной.
2. **PsExec** — fallback, если WinRM-порт 5985 закрыт ИЛИ WinRM-вызов упал по сетевой причине.
3. **MSI fallback** через WinRM (`Copy-Item -ToSession` + msiexec) для хостов без Capability API.

Категории: `INSTALLED`, `INSTALLED_VIA_PSEXEC`, `INSTALLED_MSI`, `REBOOT_REQUIRED`, `NO_CAPABILITY_API`, `SMB_UNREACHABLE`, `WINRM_CONNECT_FAIL`, `WINRM_INVOKE_FAIL`, `PSEXEC_ERROR`, `SSHD_NOT_RUNNING`.

### Deploy-SshKey (SSH-first)
1. **SSH** (Posh-SSH, парольная авторизация) — если TCP 22 открыт и пароль подходит. Post-check входа по приватному ключу.
2. **WinRM** (`Invoke-Command` с тем же scriptblock) — fallback:
   - SSH-порт закрыт → ключ кладётся **авансом** (на будущее, когда поднимут sshd)
   - SSH-порт открыт, но пароль не подходит (pubkey-only auth) → WinRM кладёт ключ, после чего SSH начнёт принимать.

Категории: `ADDED`, `SKIP_EXISTS`, `SSH_AUTH_FAIL`, `WINRM_CONNECT_FAIL`, `WINRM_INVOKE_FAIL`, `NO_TRANSPORT`.

## Параметры всех скриптов

- `-DryRun` — план без изменений
- `-HostFilter <glob>` — фильтр по имени или IP
- `-HostsFile <path>` — список хостов в файле (приоритетнее `-HostFilter`)
- `-ConfigPath <path>` — переопределить config.yaml
- `-Force` — игнорировать сохранённый state, переобработать всех

## Грабли

- ⚠️ **PowerShell parser ломается на `$role:`** — `:` после `$var` PS воспринимает как scope-prefix (`$env:`, `$global:`). Если нужно интерполировать, потом двоеточие — пиши `${role}:`. Эта ошибка уже дважды стоила времени.
- ⚠️ **`$Host` — зарезервированная read-only переменная PS** (контейнер для текущего host application). Нельзя использовать как имя параметра функции — runtime ругается `Не удается перезаписать переменную Host`. Безопасные синонимы: `$Target`, `$HostInfo`, `$Computer`.
- ⚠️ **UTF-8 BOM на .ps1 с кириллицей** обязателен. После Write всегда конвертировать через `[System.IO.File]::WriteAllText` с `UTF8Encoding($true)`.
- ⚠️ **TCP 445 (SMB) у клиента может быть закрыт** даже когда WinRM работает (политика AV/EDR). Не считать PsExec ненадёжным — это сетевое ограничение. Pre-check TCP 445 экономит таймауты.
- ⚠️ **`Add-WindowsCapability` может вернуть успех + `RestartNeeded=true`** — без перезагрузки `C:\ProgramData\ssh` не создаётся. В скрипте проверяем явно и возвращаем `REBOOT_REQUIRED`.
- ⚠️ **Set-Content / Set-Service / Restart-Service без `-ErrorAction Stop`** — non-terminating ошибки, `try/catch` их не ловит. Скрипт прогонит мимо них с ложными маркерами. Везде ставить `-ErrorAction Stop`.
- ⚠️ **`Get-LocalGroupMember -SID 'S-1-5-32-544'` на DC отдаёт `Группа S-1-5-32-544 не найдена`** — известный баг командлета, проявляется при наличии нерезолвимых SIDов среди членов группы. Фоллбек: `[ADSI]"WinNT://./Administrators,group"` + `psbase.Invoke('Members')` + чтение `objectSid` через рефлексию. Локализованное имя группы — через `Translate(NTAccount)` от well-known SID (`Администраторы` на ru-винде).
- ⚠️ **Перечисление членов BUILTIN\Administrators не видит вложенные группы**. На домен-машинах доменный админ обычно числится в локальной `BUILTIN\Administrators` через цепочку `Domain Admins → BUILTIN\Administrators` — прямого членства нет, и `Get-LocalGroupMember` / ADSI-Members его не покажут. Для проверки админства того же юзера, под которым залогинены — использовать access token: `(New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)`. Это правильный путь для Deploy-SshKey, потому что иначе ключ кладётся в `~/.ssh/authorized_keys`, а sshd для админа читает только `C:\ProgramData\ssh\administrators_authorized_keys` (default `Match Group administrators`) — и post-check фейлится с `Permission denied (publickey)`.
- ⚠️ **Inline `#`-комментарий в YAML** — мини-парсер должен срезать `#` после whitespace, иначе значение типа `password: foo  # hint` приходит как `foo  # hint` и ломает сравнения. Реализовано в `Parse-SimpleYaml`.
- ⚠️ **Пустой скаляр в YAML (`key:` без значения)** мини-парсер трактует как пустой `@{}`. `-not @{}` это `$false`, поэтому проверки вида `if (-not $val)` не срабатывают. Нормализуй через helper типа `_AsString` (hashtable → `$null`).
- ⚠️ **WinRM-клиент в рабочей группе требует настройки**: `TrustedHosts=*`, `Auth\Basic=true`, `AllowUnencrypted=true`, плюс запущенная служба `WinRM`. Иначе `Invoke-Command`/`New-PSSession` падает с `Клиенту WinRM не удается обработать запрос... TrustedHosts...`. Для правки нужны админ-права — в скриптах сделан `Assert-Admin` с UAC-elevation, плюс `Initialize-LocalWSMan`.

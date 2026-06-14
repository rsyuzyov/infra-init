# winrm-deploy

## Факты

- Скрипт `winrm-deploy/Deploy-WinRM.ps1` массово настраивает WinRM на Windows-хостах через `PsExec.exe`.
- Парсинг инвентаря — общая логика, см. [inventory-parsing.md](inventory-parsing.md).
- PsExec ходит на хост под учёткой из `config.yaml` (`credentials.username` + `credentials.password`), выполняет `Enable-PSRemoting`, включает Basic-auth, разрешает unencrypted, открывает порт 5985 в firewall.
- Путь к PsExec задаётся в `config.yaml` через `psexec_path`.

## Грабли

- ⚠️ `psexec_path` в `config.yaml` может быть относительным (`..\tools\PsExec.exe`) — тогда работает только при запуске из каталога скрипта. Лучше использовать абсолютный путь или резолвить от `$ScriptDir` в скрипте.
- ⚠️ «Отказано в доступе» (Access is denied, exitcode 5) под **локальной** учёткой-админом — это UAC remote token filtering: при удалённом входе админ получает урезанный токен без полных прав. Бьёт **оба транспорта**: PsExec — отказ на подключении к admin$/SCM; WinRM — сессия создаётся (для этого админ-прав не надо), но команды установки **внутри** сессии (`Add-WindowsCapability`, запись в HKLM, firewall) падают с «Отказано в доступе» → `WINRM_FAIL`, без fallback на PsExec (он упрётся в то же). Лечится на целевом хосте: `reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v LocalAccountTokenFilterPolicy /t REG_DWORD /d 1 /f` (сразу, без перезагрузки). Не нужно для доменных учёток и встроенного `Administrator`.
- ⚠️ PsExec пишет stderr в **OEM (cp866)**. `Get-Content` по умолчанию читает как ANSI (cp1251) → кракозябры в логе (напр. `╬Єърчрэю т фюёЄєях`). Решение: читать `-Encoding Oem`. При **ручном** запуске PsExec прямо в консоли кракозябры из той же причины (консоль ≠ cp866) — это не лечится скриптом; для читаемого вывода в сессии: `[Console]::OutputEncoding = [Text.Encoding]::GetEncoding(866)`.
- ⚠️ `Couldn't access <host>:` — это лишь **префикс** PsExec над ошибкой подключения к admin$, НЕ синоним access denied. Причину даёт вторая строка: `Access is denied`/`Отказано в доступе` (UAC token filtering — наш случай), `сетевой путь не найден` (закрыт 445), `logon failure` (неверный пароль). Поэтому детект access denied цепляем за код 5 (`ERROR_ACCESS_DENIED`) и текст `Access is denied`/`Отказано в доступе`, а не за `Couldn't access` — иначе подсказка про `LocalAccountTokenFilterPolicy` ложно сработает на сетевых проблемах и неверном пароле.

## Хелперы Common.ps1 для access denied

- `Test-AccessDenied -Message <текст> -ExitCode <код>` — детектит отказ доступа (текст EN/RU `Access is denied`/`Отказано в доступе` + код 5). Общий для PsExec stderr и WinRM-исключения.
- `Write-AccessDeniedHint -ComputerName <хост>` — печатает справку про `LocalAccountTokenFilterPolicy`.
- Подключены при `PSEXEC_FAIL` (`Deploy-WinRM.ps1`, `Deploy-OpenSSH.ps1`) и при `WINRM_FAIL` (`Deploy-OpenSSH.ps1`) — при отказе доступа в лог выводится готовый reg-однострочник.

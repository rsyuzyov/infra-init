# winrm-deploy

## Факты

- Скрипт `winrm-deploy/Deploy-WinRM.ps1` массово настраивает WinRM на Windows-хостах через `PsExec.exe`.
- Парсинг инвентаря — общая логика, см. [inventory-parsing.md](inventory-parsing.md).
- PsExec ходит на хост под учёткой из `config.yaml` (`credentials.username` + `credentials.password`), выполняет `Enable-PSRemoting`, включает Basic-auth, разрешает unencrypted, открывает порт 5985 в firewall.
- Путь к PsExec задаётся в `config.yaml` через `psexec_path`.

## Грабли

- ⚠️ `psexec_path` в `config.yaml` может быть относительным (`..\tools\PsExec.exe`) — тогда работает только при запуске из каталога скрипта. Лучше использовать абсолютный путь или резолвить от `$ScriptDir` в скрипте.

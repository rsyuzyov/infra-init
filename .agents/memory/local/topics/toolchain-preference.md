# Toolchain в infra-init

## Правило для этого проекта

В `infra-init` **только PowerShell** (.ps1). Python в этом проекте не используем — даже параллельно, даже как черновик.

**Why:** уточнено пользователем 2026-05-14. Изначально я предложил держать Python-версию рядом с PS как референс; пользователь явно удалил Python-файлы и сказал «пусть всё будет только на ps в этом проекте». Скорее всего цель — единообразие и минимум зависимостей на стороне запуска (Windows-админ без uv/python).

**How to apply:**

- Все скрипты в `infra-init` — на PowerShell, в корне проекта (`Deploy-WinRM.ps1`, `Deploy-OpenSSH.ps1`, `Deploy-SshKey.ps1`).
- Python (paramiko, pyyaml, uv) — не использовать, даже если задача удобнее ложится на Python.
- В других проектах правило не действует автоматически — там по контексту.

## Грабли PowerShell, на которые потрачено время — учесть в новых .ps1

- ⚠️ **UTF-8 BOM обязателен** для .ps1 с кириллицей. PS 5.1 без BOM читает файл как cp1251 → кириллица превращается в крокозябру, парсер ломается на двойных кавычках в строках с кириллицей. Симптом: `[Parser]::ParseFile` показывает ошибки на строках, не имеющих отношения к проблеме, так как cp1251-длина кириллицы отличается от UTF-8 и сбивает смещения. `[scriptblock]::Create($text)` показывает правильную ошибку.
- ⚠️ **Сразу после `Write` любого .ps1** — дописать BOM:
  ```powershell
  $text = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
  $utf8bom = New-Object System.Text.UTF8Encoding($true)
  [System.IO.File]::WriteAllText($path, $text, $utf8bom)
  ```
  Можно вообще писать .ps1 на латинице (русский только в комментариях логов), чтобы не зависеть от BOM.
- ⚠️ **`$role:` ломает парсер** — PowerShell видит `$role:` как scope-prefix переменной (как `$env:`, `$global:`). Если нужно интерполировать переменную перед двоеточием — пиши `${role}:`.
- ⚠️ Интерполирующий here-string `@"..."@` с большим количеством `` ` ``-эскейпов плохо читается и легко ломается. Использовать литеральный `@'...'@` + `.Replace('__PLACEHOLDER__', $value)`.
- ⚠️ Передача удалённой PowerShell-команды через SSH/`Invoke-SSHCommand`: собрать скрипт текстом, кодировать `[System.Text.Encoding]::Unicode.GetBytes` → base64, слать `powershell.exe -EncodedCommand <base64>`. Никаких проблем с кавычками и `DefaultShellCommandOption`.
- ⚠️ SSH с паролем из PowerShell — нативный `ssh.exe` не умеет читать пароль из stdin. Решение — модуль `Posh-SSH` (`New-SSHSession -Credential`, `Invoke-SSHCommand`). Требует `Install-Module Posh-SSH -Scope CurrentUser`.

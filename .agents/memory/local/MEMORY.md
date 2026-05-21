# Memory Index

Последняя оптимизация: 2026-05-13

## Проект

- `infra-init` — инфраструктурные скрипты для подготовки Windows-хостов. Все скрипты лежат в корне:
  - `Deploy-WinRM.ps1` — массовая настройка WinRM через PsExec
  - `Deploy-OpenSSH.ps1` — массовая установка OpenSSH (с MSI-фоллбеком для старых Windows)
  - `Deploy-SshKey.ps1` — массовое прописывание SSH-ключа (через Posh-SSH)
- Общий `config.yaml` (один на все скрипты), общие `tools/PsExec.exe` и `tools/OpenSSH-Win64.msi`.
- Только PowerShell, без Python — см. [topics/toolchain-preference.md](topics/toolchain-preference.md).
- Источник инвентаря — соседний проект `net-conf-gen`, output: `~\repo\net-conf-gen\output\inventory.yaml`.

## Факты

- Инвентарь от `net-conf-gen` имеет вложенную структуру:
  `all.children.managed.children.{linux,windows,devices_ssh}.children.<sub>.hosts`
  Между `all.children` и группами окружения есть прослойка `managed`.

## Грабли и уроки

- ⚠️ Парсеры inventory должны искать целевую группу (`windows`, `linux`) рекурсивно,
  а не по жёсткому пути `all.children.<group>` — иначе при добавлении промежуточных
  уровней (например, `managed`) получишь 0 хостов без явной ошибки.

## Ссылки

- [feedback_tmp_files.md](feedback_tmp_files.md) — **временные файлы — только в `tmp/`, после использования удалять**
- [topics/architecture.md](topics/architecture.md) — структура `infra-init`, общая lib, state machine `output/<domain>/<transport>/<host>.json`, транспортная стратегия для каждого скрипта
- [topics/toolchain-preference.md](topics/toolchain-preference.md) — в `infra-init` только PowerShell, без Python
- [topics/inventory-parsing.md](topics/inventory-parsing.md) — структура инвентаря, мини-парсер YAML, общие функции рекурсивного сбора хостов
- [topics/winrm-deploy.md](topics/winrm-deploy.md) — массовая настройка WinRM через PsExec
- [topics/windows-openssh.md](topics/windows-openssh.md) — authorized_keys для админов/юзеров, ACL, поведение sshd

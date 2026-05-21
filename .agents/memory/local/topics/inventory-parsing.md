# Inventory parsing (net-conf-gen → infra-init)

## Факты

- Источник инвентаря — `~\repo\net-conf-gen\output\inventory.yaml`.
- PowerShell-скрипты в `infra-init` используют собственный мини-парсер YAML (`Parse-SimpleYaml`) — без `powershell-yaml`. Поддерживает только mapping (key/value) и вложенность по отступам.
- Логика `Get-WindowsHosts` (рекурсивный поиск группы `windows` + сбор хостов) **продублирована** в нескольких скриптах: сейчас в `winrm-deploy/Deploy-WinRM.ps1` и `openssh-deploy/Deploy-OpenSSH.ps1`. Правки нужно делать в обоих.

## Структура инвентаря

```
all
└─ children
   └─ managed                ← промежуточный уровень!
      └─ children
         ├─ linux
         │   └─ children
         │      └─ linux_servers_ssh
         │         └─ hosts: { hostname: {ansible_host: ip, ...} }
         ├─ windows
         │   └─ children
         │      └─ windows_winrm
         │         └─ hosts: { ... }
         └─ devices_ssh
            └─ hosts: { ... }
```

## Грабли

- ⚠️ Группа `windows` лежит **не** в `all.children.windows`, а под промежуточным узлом `managed`: `all.children.managed.children.windows`. Жёсткий путь даёт 0 хостов без ошибки. Парсер должен искать группу рекурсивно по имени.
- ⚠️ Хосты могут быть как непосредственно в группе (`group.hosts`), так и в её подгруппах (`group.children.<sub>.hosts`). Сборщик должен рекурсивно обходить и `hosts`, и `children`.

## Рецепт

В скриптах для рекурсивного сбора используются две функции:
- `Find-GroupByName -Node <hashtable> -Name <name>` — рекурсивно ищет узел группы по имени, спускаясь в `children`.
- `Collect-HostsRecursive -Group <hashtable> -Acc <ArrayList>` — рекурсивно собирает хосты из группы и всех её детей.

См. [winrm-deploy.md](winrm-deploy.md), [windows-openssh.md](windows-openssh.md).

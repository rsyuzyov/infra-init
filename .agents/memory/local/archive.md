# Архив

## 2026-05

- [x] Починить парсинг Windows-хостов в Deploy-WinRM.ps1 (хостов было 0)
      created: 2026-05-13
      completed: 2026-05-13
      result: Get-WindowsHosts переписан на рекурсивный поиск группы 'windows'. Теперь видит 121 хост из inventory net-conf-gen

- [x] Починить парсинг Windows-хостов в Deploy-OpenSSH.ps1 (та же проблема)
      created: 2026-05-13
      completed: 2026-05-13
      result: тот же фикс, видит 121 хост

- [x] Написать скрипт массового прописывания SSH-ключа на Windows-хостах
      created: 2026-05-13
      completed: 2026-05-13
      result: две реализации рядом, общий config.yaml.example:
        - `ssh-key-deploy/Deploy-SshKey.ps1` (PowerShell, требует Posh-SSH; основной вариант)
        - `ssh-key-deploy/deploy_ssh_key.py` (Python, paramiko+pyyaml, `uv run`)
        Транспорт обоих — SSH с парольной авторизацией (после Deploy-OpenSSH). На хосте через `powershell -EncodedCommand` определяет роль юзера (admin/user) → кладёт ключ в правильный файл, выставляет ACL, идемпотентен по телу ключа.

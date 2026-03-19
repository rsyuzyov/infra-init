Групповой деплой winrm, ssh на windows хосты.  
Сначала сканируем хосты с помощью https://github.com/rsyuzyov/net-conf-gen  
После сканирования в каталоге output будет inventory.yaml - прописываем его в наш config.yaml  
Далее запускаем скрипт Deploy-WinRM.ps1 или Deploy-OpenSSH.ps1

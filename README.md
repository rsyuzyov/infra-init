Групповой деплой winrm, ssh на windows хосты.  
Сначала сканируем хосты с помощью https://github.com/rsyuzyov/net-conf-gen  
После сканирования в каталоге output будет inventory.yaml - прописываем его в наш config.yaml  
Далее запускаем Deploy-WinRM.ps1  
Затем Deploy-OpenSSH.ps1  
Для ssh затем можно прописать ключ с помощью Deploy-SshKey.ps1

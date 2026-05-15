<#
.SYNOPSIS
    Одноразовая настройка локальной машины для запуска Deploy-*.ps1 скриптов из infra-init.
.DESCRIPTION
    Что делает (требует прав админа):
      - Запускает службу WinRM.
      - Включает Auth\Basic, AllowUnencrypted в локальном WSMan-клиенте.
      - Ставит TrustedHosts=*.

    После этого Deploy-WinRM/Deploy-OpenSSH/Deploy-SshKey можно запускать
    без повышения привилегий (они сами проверяют через Test-WSManClientReady
    и не дёргают UAC, если всё уже настроено).

.EXAMPLE
    .\Setup-LocalEnvironment.ps1
#>

param()

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $ScriptDir "Common.ps1")

Assert-Admin -ScriptPath $MyInvocation.MyCommand.Definition -BoundParams $PSBoundParameters `
    -Reason "настройка локального WSMan-клиента"

Write-Host "=== Setup-LocalEnvironment ===" -ForegroundColor Cyan
Write-Host ""

if (Test-WSManClientReady) {
    Write-Host "Локальный WSMan-клиент уже настроен." -ForegroundColor Green
    Write-Host "  - Служба WinRM: запущена"
    Write-Host "  - TrustedHosts: *"
    Write-Host "  - Auth\Basic: true"
    Write-Host "  - AllowUnencrypted: true"
    Write-Host ""
    Write-Host "Никаких действий не требуется." -ForegroundColor Green
    exit 0
}

Write-Host "Настраиваю локальный WSMan-клиент..."
if (Initialize-LocalWSMan) {
    if (Test-WSManClientReady) {
        Write-Host "Готово:" -ForegroundColor Green
        Write-Host "  - Служба WinRM запущена"
        Write-Host "  - TrustedHosts=*"
        Write-Host "  - Auth\Basic=true"
        Write-Host "  - AllowUnencrypted=true"
        Write-Host ""
        Write-Host "Теперь Deploy-*.ps1 можно запускать без админа." -ForegroundColor Green
        exit 0
    } else {
        Write-Host "Initialize-LocalWSMan вернул успех, но Test-WSManClientReady всё ещё false." -ForegroundColor Yellow
        Write-Host "Проверь вручную: Get-Item WSMan:\localhost\Client\*"
        exit 2
    }
}

Write-Host "Initialize-LocalWSMan упал." -ForegroundColor Red
exit 1
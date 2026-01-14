Write-Host "Загрузка установщика Microsoft Office..." -ForegroundColor Cyan

$scriptUrl = "https://raw.githubusercontent.com/xLeizuLqrd/OfficeSelect/main/install.ps1"

try {
    Invoke-Expression (Invoke-RestMethod -Uri $scriptUrl)
} catch {
    Write-Host "Ошибка при загрузке скрипта: $_" -ForegroundColor Red
    Write-Host "Проверьте подключение к интернету и доступность репозитория." -ForegroundColor Yellow
    Pause
}

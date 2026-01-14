Write-Host "Загрузка установщика Microsoft Office..." -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Admin)) {
    Write-Host "Требуются права администратора" -ForegroundColor Yellow
    Write-Host "Перезапуск с повышенными привилегиями..." -ForegroundColor Yellow
    Write-Host ""
    
    $scriptContent = @'
$installerUrl = "https://raw.githubusercontent.com/xLeizuLqrd/OfficeSelect/main/install.ps1"
try {
    $script = Invoke-RestMethod -Uri $installerUrl
    Invoke-Expression $script
} catch {
    Write-Host "Ошибка при загрузке скрипта: $_" -ForegroundColor Red
    pause
}
'@
    
    $tempFile = [System.IO.Path]::GetTempFileName() + ".ps1"
    $scriptContent | Out-File -FilePath $tempFile -Encoding UTF8
    
    Start-Process PowerShell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$tempFile`"" -Verb RunAs
    exit
}

Write-Host "Загрузка основного установщика..." -ForegroundColor Green
Write-Host ""

try {
    $installerUrl = "https://raw.githubusercontent.com/xLeizuLqrd/OfficeSelect/main/install.ps1"
    $scriptContent = Invoke-RestMethod -Uri $installerUrl
    Invoke-Expression $scriptContent
} catch {
    Write-Host "Ошибка при загрузке скрипта: $_" -ForegroundColor Red
    Write-Host "Проверьте подключение к интернету." -ForegroundColor Yellow
    pause
}

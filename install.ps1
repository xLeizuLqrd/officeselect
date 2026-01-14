[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "Установщик Microsoft Office"

function Show-ModeMenu {
    Clear-Host
    Write-Host @"
        :::    :::       :::        ::::::::::       :::::::::::       :::::::::      :::    :::
       :+:    :+:       :+:        :+:                  :+:                :+:       :+:    :+:
       +:+  +:+        +:+        +:+                  +:+               +:+        +:+    +:+ 
       +#++:+         +#+        +#++:++#             +#+              +#+         +#+    +:+  
     +#+  +#+        +#+        +#+                  +#+             +#+          +#+    +#+   
   #+#    #+#       #+#        #+#                  #+#            #+#           #+#    #+#    
  ###    ###       ########## ##########       ###########       #########       ########      

========================================
     ВЫБОР РЕЖИМА УСТАНОВКИ
========================================

[1] Полная установка (удалит старый Office)
[2] Добавить программы к существующему Office

========================================
"@ -ForegroundColor Cyan
    
    do {
        $mode = Read-Host "`nВыберите режим (1 или 2)"
        
        switch ($mode) {
            "1" {
                Clear-Host
                Write-Host "========================================" -ForegroundColor Cyan
                Write-Host "ВЫБРАН РЕЖИМ: ПОЛНАЯ УСТАНОВКА" -ForegroundColor Yellow
                Write-Host "========================================" -ForegroundColor Cyan
                Write-Host ""
                Write-Host "ВНИМАНИЕ: Существующий Office будет удален!" -ForegroundColor Red
                Write-Host ""
                Start-Sleep -Seconds 2
                $script:RemoveMSI = $true
                $script:ModeName = "ПОЛНАЯ УСТАНОВКА"
                Show-MainMenu
                return
            }
            "2" {
                Clear-Host
                Write-Host "========================================" -ForegroundColor Cyan
                Write-Host "ВЫБРАН РЕЖИМ: ДОБАВЛЕНИЕ ПРОГРАММ" -ForegroundColor Yellow
                Write-Host "========================================" -ForegroundColor Cyan
                Write-Host ""
                Write-Host "Office будет установлен поверх существующего."
                Write-Host ""
                Start-Sleep -Seconds 2
                $script:RemoveMSI = $false
                $script:ModeName = "ДОБАВЛЕНИЕ ПРОГРАММ"
                Show-MainMenu
                return
            }
            default {
                Write-Host "`nОшибка: Неверный выбор! Введите 1 или 2" -ForegroundColor Red
            }
        }
    } while ($true)
}

function Show-MainMenu {
    Clear-Host
    Write-Host @"
========================================
     ВЫБОР ПРОГРАММ MICROSOFT OFFICE
========================================

Выберите цифры через пробел:

[1] Word       [2] Excel
[3] PowerPoint [4] Outlook
[5] Access     [6] Publisher
[7] OneNote    [8] OneDrive
[9] Teams      [10] Groove
[11] Lync      [0] Назад

Просто Enter - установить все программы
Перед установкой убедитесь, что все программы Office закрыты и данные сохранены!
РЕЖИМ: $($script:ModeName)

========================================
"@ -ForegroundColor Cyan
    
    $input = Read-Host "`nВаш выбор"
    
    if ($input -eq "0") {
        Show-ModeMenu
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($input)) {
        Clear-Host
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "ВЫБРАНЫ ВСЕ ПРОГРАММЫ" -ForegroundColor Yellow
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Будет установлено все 11 программ Office"
        Write-Host ""
        Start-Sleep -Seconds 2
        $script:InstallAll = $true
        $script:SelectedApps = @(1..11)
    } else {
        if ($input -match '^[0-9\s]+$') {
            $script:InstallAll = $false
            $script:SelectedApps = $input -split '\s+' | ForEach-Object { [int]$_ } | Where-Object { $_ -ge 1 -and $_ -le 11 }
            
            if ($script:SelectedApps.Count -eq 0) {
                Write-Host "`nОшибка: Неверный ввод!" -ForegroundColor Red
                Write-Host "Вводите только цифры от 1 до 11 через пробел!" -ForegroundColor Yellow
                Start-Sleep -Seconds 3
                Show-MainMenu
                return
            }
        } else {
            Write-Host "`nОшибка: Неверный формат ввода!" -ForegroundColor Red
            Write-Host "Вводите только цифры от 1 до 11 через пробел!" -ForegroundColor Yellow
            Start-Sleep -Seconds 3
            Show-MainMenu
            return
        }
    }
    
    Create-Configuration
}

function Create-Configuration {
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "     СОЗДАНИЕ КОНФИГУРАЦИИ" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        $script:RegBackup = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" -Name "CountryCode" -ErrorAction SilentlyContinue
    } catch {
        $script:RegBackup = $null
    }
    
    Write-Host "Изменение настроек реестра..." -ForegroundColor Gray
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" -Name "CountryCode" -Value "std::wstring|US" -Force -ErrorAction SilentlyContinue
    Write-Host "Готово!" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "Создание configuration.xml..." -ForegroundColor Gray
    
    $xmlContent = @()
    $xmlContent += '<?xml version="1.0" encoding="utf-8"?>'
    $xmlContent += '<Configuration>'
    $xmlContent += '  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">'
    $xmlContent += '    <Product ID="ProPlus2024Volume" PIDKEY="Y63J7-9RNDJ-GD3BV-BDKBP-HH966">'
    $xmlContent += '      <Language ID="ru-ru" />'
    
    if ($script:InstallAll) {
        $xmlContent += '      <!-- Все программы будут установлены -->'
        Clear-Host
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "     УСТАНОВКА ВСЕХ ПРОГРАММ" -ForegroundColor Yellow
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "БУДУТ УСТАНОВЛЕНЫ ВСЕ 11 ПРОГРАММ OFFICE" -ForegroundColor Green
        Write-Host ""
        Start-Sleep -Seconds 2
    } else {
        $appMap = @{
            1 = "Word"
            2 = "Excel"
            3 = "PowerPoint"
            4 = "Outlook"
            5 = "Access"
            6 = "Publisher"
            7 = "OneNote"
            8 = "OneDrive"
            9 = "Teams"
            10 = "Groove"
            11 = "Lync"
        }
        
        $excludeApps = @()
        $includeApps = @()
        
        $allApps = 1..11
        
        foreach ($appNum in $allApps) {
            if ($appNum -notin $script:SelectedApps) {
                $excludeApps += $appMap[$appNum]
            } else {
                $includeApps += $appMap[$appNum]
            }
        }
        
        foreach ($app in $excludeApps) {
            $xmlContent += "      <ExcludeApp ID=`"$app`" />"
        }
        
        Clear-Host
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "     ВЫБОР ПРОГРАММ" -ForegroundColor Yellow
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        
        if ($includeApps.Count -gt 0) {
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host "     БУДУТ УСТАНОВЛЕНЫ:" -ForegroundColor Green
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host ($includeApps -join ", ")
            Write-Host ""
        }
        
        if ($excludeApps.Count -gt 0) {
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host "     НЕ БУДУТ УСТАНОВЛЕНЫ:" -ForegroundColor Gray
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host ($excludeApps -join ", ")
            Write-Host ""
        }
        
        Start-Sleep -Seconds 3
    }
    
    $xmlContent += '    </Product>'
    $xmlContent += '  </Add>'
    
    if ($script:RemoveMSI) {
        $xmlContent += '  <RemoveMSI />'
        Clear-Host
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "     ВНИМАНИЕ!" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Существующий Office будет УДАЛЕН!" -ForegroundColor Red
        Write-Host ""
        Start-Sleep -Seconds 3
    }
    
    $xmlContent += '  <Property Name="AUTOACTIVATE" Value="1" />'
    $xmlContent += '</Configuration>'
    
    $xmlContent | Out-File -FilePath "configuration.xml" -Encoding UTF8 -Force
    
    Write-Host "Закрытие запущенных программ Office..." -ForegroundColor Gray
    $officeProcesses = @("winword", "excel", "powerpnt", "outlook", "msaccess", "onenote", "mspub")
    foreach ($process in $officeProcesses) {
        Get-Process -Name $process -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 1
    
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "КОНФИГУРАЦИЯ СОЗДАНА" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Файл configuration.xml создан успешно!" -ForegroundColor Green
    Write-Host ""
    Start-Sleep -Seconds 2
    
    Start-Installation
}

function Start-Installation {
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "     ЗАПУСК УСТАНОВКИ OFFICE" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Установка Microsoft Office начата..." -ForegroundColor Green
    Write-Host ""
    
    Write-Host "1. Скачивание Office Deployment Tool..." -ForegroundColor Gray
    $odtUrl = "https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_19426-20170.exe"
    $odtPath = "$env:TEMP\ODTSetup.exe"
    
    try {
        Invoke-WebRequest -Uri $odtUrl -OutFile $odtPath -UseBasicParsing
        Write-Host "   ✓ ODT скачан" -ForegroundColor Green
    } catch {
        Write-Host "   ✗ Ошибка скачивания ODT" -ForegroundColor Red
        Write-Host "   Скачайте вручную: https://aka.ms/ODT" -ForegroundColor Yellow
        pause
        exit 1
    }
    
    Write-Host "2. Извлечение файлов установки..." -ForegroundColor Gray
    $extractDir = "$env:TEMP\OfficeSetup"
    Remove-Item -Path $extractDir -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path $extractDir -Force | Out-Null
    
    Start-Process -FilePath $odtPath -ArgumentList "/extract:`"$extractDir`" /quiet" -Wait -NoNewWindow
    Write-Host "   ✓ Файлы извлечены" -ForegroundColor Green
    
    Write-Host "3. Подготовка установки..." -ForegroundColor Gray
    $sourceSetup = Join-Path $extractDir "setup.exe"
    $destSetup = "setup.exe"
    
    if (Test-Path $sourceSetup) {
        Copy-Item -Path $sourceSetup -Destination $destSetup -Force
        Write-Host "   ✓ Установщик готов" -ForegroundColor Green
    } else {
        Write-Host "   ✗ Ошибка: setup.exe не найден" -ForegroundColor Red
        pause
        exit 1
    }
    
    Write-Host "4. Запуск установки Office..." -ForegroundColor Gray
    Write-Host "   Это может занять 10-30 минут" -ForegroundColor Yellow
    Write-Host "   Не закрывайте это окно!" -ForegroundColor Red
    Write-Host ""
    
    try {
        $process = Start-Process -FilePath $destSetup -ArgumentList "/configure configuration.xml" -Wait -NoNewWindow -PassThru
        
        Remove-Item -Path $odtPath -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $extractDir -Recurse -Force -ErrorAction SilentlyContinue
        
        if ($process.ExitCode -eq 0) {
            Write-Host "   ✓ Установка завершена успешно!" -ForegroundColor Green
        } else {
            Write-Host "   ⚠ Установка завершена с кодом: $($process.ExitCode)" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "   ✗ Ошибка при установке: $_" -ForegroundColor Red
    }
    
    Write-Host ""
    Write-Host "Для продолжения нажмите Enter..." -ForegroundColor Gray
    pause
    
    Restore-Settings
}

function Restore-Settings {
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "     ВОССТАНОВЛЕНИЕ НАСТРОЕК" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Восстановление настроек реестра..." -ForegroundColor Gray
    
    if ($script:RegBackup -and $script:RegBackup.CountryCode) {
        Write-Host "Восстанавливаем оригинальное значение..." -ForegroundColor Gray
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" -Name "CountryCode" -Value $script:RegBackup.CountryCode -Force -ErrorAction SilentlyContinue
        Write-Host "Готово!" -ForegroundColor Green
    } else {
        Write-Host "Удаляем временный параметр..." -ForegroundColor Gray
        Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" -Name "CountryCode" -ErrorAction SilentlyContinue
        Write-Host "Готово!" -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "Удаление временных файлов..." -ForegroundColor Gray
    Remove-Item -Path "configuration.xml" -ErrorAction SilentlyContinue
    Write-Host "Готово!" -ForegroundColor Green
    Write-Host ""
    Start-Sleep -Seconds 2
    
    Show-FinishMenu
}

function Show-FinishMenu {
    Clear-Host
    Write-Host @"
========================================
     УСТАНОВКА ЗАВЕРШЕНА
========================================

Для начала работы откройте любое приложение Office.

========================================

[1] Вернуться в главное меню
[2] Выход

========================================
"@ -ForegroundColor Green
    
    do {
        $choice = Read-Host "`nВаш выбор"
        
        switch ($choice) {
            "1" {
                Show-ModeMenu
                return
            }
            "2" {
                exit
            }
            default {
                Write-Host "Неверный выбор. Введите 1 или 2" -ForegroundColor Red
            }
        }
    } while ($true)
}

$script:RemoveMSI = $false
$script:ModeName = ""
$script:InstallAll = $false
$script:SelectedApps = @()
$script:RegBackup = $null

Show-ModeMenu

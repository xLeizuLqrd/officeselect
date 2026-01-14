function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Start-AsAdmin {
    if (-not (Test-Admin)) {
        Write-Host "Запуск от имени администратора..." -ForegroundColor Yellow
        $scriptPath = $MyInvocation.MyCommand.Path
        Start-Process PowerShell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
        exit
    }
}

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
    
    $mode = Read-Host "Выберите режим (1 или 2)"
    
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
        }
        default {
            Write-Host "Ошибка: Неверный выбор!" -ForegroundColor Red
            Write-Host "Пожалуйста, введите 1 или 2" -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            Show-ModeMenu
        }
    }
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
    
    $input = Read-Host "Ваш выбор"
    
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
                Write-Host "Ошибка: Неверный ввод!" -ForegroundColor Red
                Write-Host "Вводите только цифры от 1 до 11 через пробел!" -ForegroundColor Yellow
                Start-Sleep -Seconds 3
                Show-MainMenu
                return
            }
        } else {
            Write-Host "Ошибка: Неверный формат ввода!" -ForegroundColor Red
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
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" -Name "CountryCode" -Value "std::wstring|US" -Force
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
    
    $xmlContent | Out-File -FilePath "configuration.xml" -Encoding UTF8
    
    Write-Host "Закрытие запущенных программ Office..." -ForegroundColor Gray
    $officeProcesses = @("winword", "excel", "powerpnt", "outlook", "msaccess", "onenote", "mspub")
    foreach ($process in $officeProcesses) {
        Get-Process -Name $process -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    
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
    Write-Host "Это может занять несколько минут." -ForegroundColor Yellow
    Write-Host "Пожалуйста, не закрывайте окно." -ForegroundColor Yellow
    Write-Host ""
    
    try {
        $setupPath = Join-Path $PSScriptRoot "setup.exe"
        
        if (Test-Path $setupPath) {
            Start-Process -FilePath $setupPath -ArgumentList "/configure configuration.xml" -Wait -NoNewWindow
        } else {
            Write-Host "Ошибка: setup.exe не найден!" -ForegroundColor Red
            Write-Host "Поместите setup.exe в ту же папку, что и скрипт." -ForegroundColor Yellow
            Pause
            exit 1
        }
    } catch {
        Write-Host "Ошибка при установке: $_" -ForegroundColor Red
        Pause
        exit 1
    }
    
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
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" -Name "CountryCode" -Value $script:RegBackup.CountryCode -Force
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
     УСТАНОВКА УСПЕШНО ЗАВЕРШЕНА!
========================================

✓ Microsoft Office успешно установлен!

Для начала работы откройте любое приложение Office.

========================================

[1] Вернуться в главное меню
[2] Выход

========================================
"@ -ForegroundColor Green
    
    $choice = Read-Host "Ваш выбор"
    
    switch ($choice) {
        "1" {
            Show-ModeMenu
        }
        "2" {
            exit
        }
        default {
            Write-Host "Ошибка: Неверный выбор!" -ForegroundColor Red
            Write-Host "Пожалуйста, введите 1 или 2" -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            Show-FinishMenu
        }
    }
}

function Main {
    Start-AsAdmin
    
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    
    $script:RemoveMSI = $false
    $script:ModeName = ""
    $script:InstallAll = $false
    $script:SelectedApps = @()
    $script:RegBackup = $null
    
    Show-ModeMenu
}

Main

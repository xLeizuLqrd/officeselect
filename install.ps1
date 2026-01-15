[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "Установщик Microsoft Office"

function Show-ModeMenu {
    Clear-Host
    Write-Host "        :::    :::       :::        ::::::::::       :::::::::::       :::::::::      :::    :::" -ForegroundColor Cyan
    Write-Host "       :+:    :+:       :+:        :+:                  :+:                :+:       :+:    :+:" -ForegroundColor Cyan
    Write-Host "       +:+  +:+        +:+        +:+                  +:+               +:+        +:+    +:+ " -ForegroundColor Cyan
    Write-Host "       +#++:+         +#+        +#++:++#             +#+              +#+         +#+    +:+  " -ForegroundColor Cyan
    Write-Host "     +#+  +#+        +#+        +#+                  +#+             +#+          +#+    +#+   " -ForegroundColor Cyan
    Write-Host "   #+#    #+#       #+#        #+#                  #+#            #+#           #+#    #+#    " -ForegroundColor Cyan
    Write-Host "  ###    ###       ########## ##########       ###########       #########       ########      " -ForegroundColor Cyan
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "     ВЫБОР РЕЖИМА УСТАНОВКИ" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "[1] Полная установка (удалит старый Office)" -ForegroundColor Yellow
    Write-Host "[2] Добавить программы к существующему Office" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
        
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
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" -Name "CountryCode" -Value "US" -Force -ErrorAction SilentlyContinue
    Write-Host "Готово!" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "Создание configuration.xml..." -ForegroundColor Gray
    
    $xmlContent = @()
    $xmlContent += '<?xml version="1.0" encoding="utf-8"?>'
    $xmlContent += '<Configuration>'
    $xmlContent += '  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">'
    $xmlContent += '    <Product ID="ProPlus2024Volume">'
    $xmlContent += '      <Language ID="ru-ru" />'
    
    if (-not $script:InstallAll) {
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
        
        $allApps = 1..11
        foreach ($appNum in $allApps) {
            if ($appNum -notin $script:SelectedApps) {
                $xmlContent += "      <ExcludeApp ID=`"$($appMap[$appNum])`" />"
            }
        }
    }
    
    $xmlContent += '    </Product>'
    $xmlContent += '  </Add>'
    
    if ($script:RemoveMSI) {
        $xmlContent += '  <RemoveMSI All="True" />'
    }
    
    $xmlContent += '  <Display Level="None" AcceptEULA="TRUE" />'
    $xmlContent += '  <Property Name="AUTOACTIVATE" Value="1" />'
    $xmlContent += '  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />'
    $xmlContent += '</Configuration>'
    
    $xmlContent | Out-File -FilePath "configuration.xml" -Encoding UTF8 -Force
    
    $officeProcesses = @("winword", "excel", "powerpnt", "outlook", "msaccess", "onenote", "mspub")
    foreach ($process in $officeProcesses) {
        Get-Process -Name $process -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 2
    
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
    
    $workDir = "$env:TEMP\OfficeInstall_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    New-Item -ItemType Directory -Path $workDir -Force | Out-Null
    
    try {
        Write-Host "1. Скачивание Office Deployment Tool..." -ForegroundColor Gray
        $odtUrl = "https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_19426-20170.exe"
        $odtPath = Join-Path $workDir "ODTSetup.exe"
        
        $progressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $odtUrl -OutFile $odtPath -UseBasicParsing
        $progressPreference = 'Continue'
        
        if (-not (Test-Path $odtPath)) {
            throw "Не удалось скачать ODT"
        }
        Write-Host "   ✓ ODT скачан" -ForegroundColor Green
        
        Write-Host "2. Извлечение файлов установки..." -ForegroundColor Gray
        $extractDir = Join-Path $workDir "OfficeSetup"
        New-Item -ItemType Directory -Path $extractDir -Force | Out-Null
        
        Start-Process -FilePath $odtPath -ArgumentList "/extract:`"$extractDir`" /quiet" -Wait -NoNewWindow
        
        $setupFiles = Get-ChildItem -Path $extractDir -Filter "setup.exe" -Recurse
        if ($setupFiles.Count -eq 0) {
            throw "Не удалось извлечь setup.exe"
        }
        Write-Host "   ✓ Файлы извлечены" -ForegroundColor Green
        
        Write-Host "3. Подготовка установки..." -ForegroundColor Gray
        
        $sourceSetup = $setupFiles[0].FullName
        $destSetup = Join-Path $workDir "setup.exe"
        Copy-Item -Path $sourceSetup -Destination $destSetup -Force
        
        $configSource = "configuration.xml"
        $configDest = Join-Path $workDir "configuration.xml"
        if (-not (Test-Path $configSource)) {
            throw "Не найден файл конфигурации"
        }
        Copy-Item -Path $configSource -Destination $configDest -Force
        
        Write-Host "   ✓ Установщик готов" -ForegroundColor Green
        
        Write-Host "4. Проверка конфигурации..." -ForegroundColor Gray
        $configContent = Get-Content -Path $configDest -Raw
        if ($configContent -notmatch "<Configuration>") {
            throw "Неверный формат конфигурационного файла"
        }
        Write-Host "   ✓ Конфигурация проверена" -ForegroundColor Green
        
        Write-Host "5. Запуск установки Office..." -ForegroundColor Gray
        Write-Host "   Это может занять 10-30 минут" -ForegroundColor Yellow
        Write-Host "   Не закрывайте это окно!" -ForegroundColor Red
        Write-Host ""
        
        Push-Location $workDir
        
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = "setup.exe"
        $processInfo.Arguments = "/configure configuration.xml"
        $processInfo.WorkingDirectory = $workDir
        $processInfo.UseShellExecute = $false
        $processInfo.RedirectStandardOutput = $true
        $processInfo.RedirectStandardError = $true
        $processInfo.CreateNoWindow = $false
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        
        $outputBuilder = New-Object System.Text.StringBuilder
        $errorBuilder = New-Object System.Text.StringBuilder
        
        $outputHandler = {
            if (![String]::IsNullOrEmpty($EventArgs.Data)) {
                $outputBuilder.AppendLine($EventArgs.Data)
            }
        }
        
        $errorHandler = {
            if (![String]::IsNullOrEmpty($EventArgs.Data)) {
                $errorBuilder.AppendLine($EventArgs.Data)
            }
        }
        
        $eventOutput = Register-ObjectEvent -InputObject $process -EventName 'OutputDataReceived' -Action $outputHandler
        $eventError = Register-ObjectEvent -InputObject $process -EventName 'ErrorDataReceived' -Action $errorHandler
        
        $process.Start() | Out-Null
        $process.BeginOutputReadLine()
        $process.BeginErrorReadLine()
        $process.WaitForExit()
        
        $exitCode = $process.ExitCode
        $output = $outputBuilder.ToString()
        $errorOutput = $errorBuilder.ToString()
        
        Pop-Location
        
        Write-Host "`n6. Результат установки:" -ForegroundColor Gray
        
        if ($exitCode -eq 0) {
            Write-Host "   ✓ Установка завершена успешно!" -ForegroundColor Green
        } elseif ($exitCode -eq 3010) {
            Write-Host "   ⚠ Требуется перезагрузка системы" -ForegroundColor Yellow
        } elseif ($exitCode -eq 17002) {
            Write-Host "   ⚠ Office уже установлен или обновлен" -ForegroundColor Yellow
        } else {
            Write-Host "   ✗ Ошибка установки (код: $exitCode)" -ForegroundColor Red
            
            if ($errorOutput) {
                Write-Host "`nДетали ошибки:" -ForegroundColor Red
                Write-Host $errorOutput -ForegroundColor Red
            }
            
            Write-Host "`nВозможные решения:" -ForegroundColor Yellow
            Write-Host "1. Закройте все программы Office" -ForegroundColor Yellow
            Write-Host "2. Отключите антивирус на время установки" -ForegroundColor Yellow
            Write-Host "3. Удалите старый Office перед установкой" -ForegroundColor Yellow
            Write-Host "4. Проверьте свободное место на диске" -ForegroundColor Yellow
        }
        
        Unregister-Event -SourceIdentifier $eventOutput.Name -ErrorAction SilentlyContinue
        Unregister-Event -SourceIdentifier $eventError.Name -ErrorAction SilentlyContinue
        
    } catch {
        Write-Host "`n✗ Критическая ошибка: $_" -ForegroundColor Red
        Write-Host "Для продолжения нажмите Enter..." -ForegroundColor Gray
        pause
    } finally {
        try {
            if (Test-Path $workDir) {
                Remove-Item -Path $workDir -Recurse -Force -ErrorAction SilentlyContinue
            }
        } catch {}
    }
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

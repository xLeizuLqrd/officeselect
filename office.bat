@echo off
chcp 65001
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo Запуск от имени администратора...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)
cd /d "%~dp0"
del configuration.xml >nul 2>&1
title Установка Microsoft Office
setlocal enabledelayedexpansion

if not exist "setup.exe" (
    echo Ошибка: setup.exe не найден в текущей папке!
    pause
    exit /b 1
)

:mode_menu
cls    
echo        :::    :::       :::        ::::::::::       :::::::::::       :::::::::      :::    :::
echo       :+:    :+:       :+:        :+:                  :+:                :+:       :+:    :+:
echo       +:+  +:+        +:+        +:+                  +:+               +:+        +:+    +:+ 
echo       +#++:+         +#+        +#++:++#             +#+              +#+         +#+    +:+  
echo     +#+  +#+        +#+        +#+                  +#+             +#+          +#+    +#+   
echo   #+#    #+#       #+#        #+#                  #+#            #+#           #+#    #+#    
echo  ###    ###       ########## ##########       ###########       #########       ########      
echo.
echo ========================================
echo     ВЫБОР РЕЖИМА УСТАНОВКИ
echo ========================================
echo.
echo [1] Полная установка (удалит старый Office)
echo [2] Добавить программы к существующему Office
echo.
echo ========================================
echo.

:ask_mode
set "install_mode="
set /p "install_mode=Выберите режим (1 или 2): "

if "!install_mode!"=="1" (
    cls
    echo ========================================
    echo ВЫБРАН РЕЖИМ: ПОЛНАЯ УСТАНОВКА
    echo ========================================
    echo.
    echo ВНИМАНИЕ: Существующий Office будет удален!
    echo.
    timeout /t 2 /nobreak >nul
    set "remove_msi=<RemoveMSI />"
    set "mode_name=ПОЛНАЯ УСТАНОВКА"
    goto main_menu
    
) else if "!install_mode!"=="2" (
    cls
    echo ========================================
    echo ВЫБРАН РЕЖИМ: ДОБАВЛЕНИЕ ПРОГРАММ
    echo ========================================
    echo.
    echo Office будет установлен поверх существующего.
    echo.
    timeout /t 2 /nobreak >nul
    set "remove_msi="
    set "mode_name=ДОБАВЛЕНИЕ ПРОГРАММ"
    goto main_menu
) else (
    cls
    echo ========================================
    echo ОШИБКА: НЕВЕРНЫЙ ВЫБОР!
    echo ========================================
    echo.
    echo Пожалуйста, введите 1 или 2
    echo.
    timeout /t 2 /nobreak >nul
    goto mode_menu
)

:main_menu
cls  
echo.
echo ========================================
echo     ВЫБОР ПРОГРАММ MICROSOFT OFFICE
echo ========================================
echo.
echo Выберите цифры через пробел:   
echo.
echo [1] Word       [2] Excel
echo [3] PowerPoint [4] Outlook
echo [5] Access     [6] Publisher
echo [7] OneNote    [8] OneDrive
echo [9] Teams      [10] Groove
echo [11] Lync echo [0] Назад
echo.
echo Просто Enter - установить все программы
echo Перед установкой убедитесь, что все программы Office закрыты и данные сохранены!
echo РЕЖИМ: !mode_name!
echo.
echo ========================================
echo.

:ask_input
set "user_input="
set /p "user_input=Ваш выбор: "

if "!user_input!"=="0" (
    goto mode_menu
)

if "!user_input!"=="" (
    cls
    echo ========================================
    echo ВЫБРАНЫ ВСЕ ПРОГРАММЫ
    echo ========================================
    echo.
    echo Будет установлено все 11 программ Office
    echo.
    timeout /t 2 /nobreak >nul
    set "install_all=1"
    goto create_xml
)

echo !user_input! | findstr /r "^[0-9 ]*$" >nul
if errorlevel 1 (
    cls
    echo ========================================
    echo ОШИБКА: НЕВЕРНЫЙ ФОРМАТ ВВОДА!
    echo ========================================
    echo.
    echo Вводите только цифры от 1 до 11 через пробел!
    echo Для возврата к выбору режима введите 0
    echo.
    timeout /t 3 /nobreak >nul
    goto main_menu
)

set "install_all=0"

:create_xml
cls
echo ========================================
echo     СОЗДАНИЕ КОНФИГУРАЦИИ
echo ========================================
echo.

set "reg_backup="
for /f "tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" /v "CountryCode" 2^>nul') do (
    set "reg_backup=%%b"
)

echo Изменение настроек реестра...
reg add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" /v "CountryCode" /t REG_SZ /d "std::wstring|US" /f
echo Готово!
echo.

echo Создание configuration.xml...
echo ^<Configuration^> > configuration.xml
echo   ^<Add OfficeClientEdition="64" Channel="PerpetualVL2024"^> >> configuration.xml
echo     ^<Product ID="ProPlus2024Volume" PIDKEY="Y63J7-9RNDJ-GD3BV-BDKBP-HH966"^> >> configuration.xml
echo       ^<Language ID="ru-ru" /^> >> configuration.xml

if "!install_all!"=="1" (
    echo         ^<!-- Все программы будут установлены --^> >> configuration.xml
    cls
    echo ========================================
    echo     УСТАНОВКА ВСЕХ ПРОГРАММ
    echo ========================================
    echo.
    echo БУДУТ УСТАНОВЛЕНЫ ВСЕ 11 ПРОГРАММ OFFICE
    echo.
    timeout /t 2 /nobreak >nul
) else (
    set "exclude_list="
    set "include_list="
    
    set "selected_word=0"
    set "selected_excel=0"
    set "selected_powerpoint=0"
    set "selected_outlook=0"
    set "selected_access=0"
    set "selected_publisher=0"
    set "selected_onenote=0"
    set "selected_onedrive=0"
    set "selected_teams=0"
    set "selected_groove=0"
    set "selected_lync=0"
    
    for %%i in (!user_input!) do (
        if "%%i"=="1" set "selected_word=1"
        if "%%i"=="2" set "selected_excel=1"
        if "%%i"=="3" set "selected_powerpoint=1"
        if "%%i"=="4" set "selected_outlook=1"
        if "%%i"=="5" set "selected_access=1"
        if "%%i"=="6" set "selected_publisher=1"
        if "%%i"=="7" set "selected_onenote=1"
        if "%%i"=="8" set "selected_onedrive=1"
        if "%%i"=="9" set "selected_teams=1"
        if "%%i"=="10" set "selected_groove=1"
        if "%%i"=="11" set "selected_lync=1"
    )
    
    cls
    echo ========================================
    echo     ВЫБОР ПРОГРАММ
    echo ========================================
    echo.
    
    if !selected_word!==0 (
        echo         ^<ExcludeApp ID="Word" /^> >> configuration.xml
        set "exclude_list=!exclude_list!Word, "
    ) else (
        echo   ✓ Word
        set "include_list=!include_list!Word, "
    )
    if !selected_excel!==0 (
        echo         ^<ExcludeApp ID="Excel" /^> >> configuration.xml
        set "exclude_list=!exclude_list!Excel, "
    ) else (
        echo   ✓ Excel
        set "include_list=!include_list!Excel, "
    )
    if !selected_powerpoint!==0 (
        echo         ^<ExcludeApp ID="PowerPoint" /^> >> configuration.xml
        set "exclude_list=!exclude_list!PowerPoint, "
    ) else (
        echo   ✓ PowerPoint
        set "include_list=!include_list!PowerPoint, "
    )
    if !selected_outlook!==0 (
        echo         ^<ExcludeApp ID="Outlook" /^> >> configuration.xml
        set "exclude_list=!exclude_list!Outlook, "
    ) else (
        echo   ✓ Outlook
        set "include_list=!include_list!Outlook, "
    )
    if !selected_access!==0 (
        echo         ^<ExcludeApp ID="Access" /^> >> configuration.xml
        set "exclude_list=!exclude_list!Access, "
    ) else (
        echo   ✓ Access
        set "include_list=!include_list!Access, "
    )
    if !selected_publisher!==0 (
        echo         ^<ExcludeApp ID="Publisher" /^> >> configuration.xml
        set "exclude_list=!exclude_list!Publisher, "
    ) else (
        echo   ✓ Publisher
        set "include_list=!include_list!Publisher, "
    )
    if !selected_onenote!==0 (
        echo         ^<ExcludeApp ID="OneNote" /^> >> configuration.xml
        set "exclude_list=!exclude_list!OneNote, "
    ) else (
        echo   ✓ OneNote
        set "include_list=!include_list!OneNote, "
    )
    if !selected_onedrive!==0 (
        echo         ^<ExcludeApp ID="OneDrive" /^> >> configuration.xml
        set "exclude_list=!exclude_list!OneDrive, "
    ) else (
        echo   ✓ OneDrive
        set "include_list=!include_list!OneDrive, "
    )
    if !selected_teams!==0 (
        echo         ^<ExcludeApp ID="Teams" /^> >> configuration.xml
        set "exclude_list=!exclude_list!Teams, "
    ) else (
        echo   ✓ Teams
        set "include_list=!include_list!Teams, "
    )
    if !selected_groove!==0 (
        echo         ^<ExcludeApp ID="Groove" /^> >> configuration.xml
        set "exclude_list=!exclude_list!Groove, "
    ) else (
        echo   ✓ Groove
        set "include_list=!include_list!Groove, "
    )
    if !selected_lync!==0 (
        echo         ^<ExcludeApp ID="Lync" /^> >> configuration.xml
        set "exclude_list=!exclude_list!Lync, "
    ) else (
        echo   ✓ Lync
        set "include_list=!include_list!Lync, "
    )
    
    echo.
    if "!include_list!" NEQ "" (
        echo ========================================
        echo     БУДУТ УСТАНОВЛЕНЫ:
        echo ========================================
        echo !include_list:~0,-2!
        echo.
    )
    
    if "!exclude_list!" NEQ "" (
        echo ========================================
        echo     НЕ БУДУТ УСТАНОВЛЕНЫ:
        echo ========================================
        echo !exclude_list:~0,-2!
        echo.
    )
    
    timeout /t 3 /nobreak >nul
)

echo     ^</Product^> >> configuration.xml
echo   ^</Add^> >> configuration.xml

if defined remove_msi (
    echo   !remove_msi! >> configuration.xml
    cls
    echo ========================================
    echo     ВНИМАНИЕ!
    echo ========================================
    echo.
    echo Существующий Office будет УДАЛЕН!
    echo.
    timeout /t 3 /nobreak >nul
)

echo   ^<Property Name="AUTOACTIVATE" Value="1" /^> >> configuration.xml
echo ^</Configuration^> >> configuration.xml

    taskkill /f /im winword.exe >nul 2>&1
    taskkill /f /im excel.exe >nul 2>&1
    taskkill /f /im powerpnt.exe >nul 2>&1
    taskkill /f /im outlook.exe >nul 2>&1
    taskkill /f /im msaccess.exe >nul 2>&1
    taskkill /f /im onenote.exe >nul 2>&1
    taskkill /f /im mspub.exe >nul 2>&1

cls
echo ========================================
echo КОНФИГУРАЦИЯ СОЗДАНА
echo ========================================
echo.
echo Файл configuration.xml создан успешно!
echo.
timeout /t 2 /nobreak >nul

:start_installation
cls
echo ========================================
echo     ЗАПУСК УСТАНОВКИ OFFICE
echo ========================================
echo.
echo Установка Microsoft Office начата...
echo.
echo Это может занять несколько минут.
echo Пожалуйста, не закрывайте окно.
echo.

setup.exe /configure configuration.xml

cls
echo ========================================
echo     ВОССТАНОВЛЕНИЕ НАСТРОЕК
echo ========================================
echo.
echo Восстановление настроек реестра...

if defined reg_backup (
    echo Восстанавливаем оригинальное значение...
    reg add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" /v "CountryCode" /t REG_SZ /d "!reg_backup!" /f
    echo Готово!
) else (
    echo Удаляем временный параметр...
    reg delete "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" /v "CountryCode" /f >nul 2>&1
    echo Готово!
)

echo.
echo Удаление временных файлов...
del configuration.xml >nul 2>&1
echo Готово!
echo.
timeout /t 2 /nobreak >nul

:finish
cls   
echo.
echo ========================================
echo     УСТАНОВКА УСПЕШНО ЗАВЕРШЕНА!
echo ========================================
echo.
echo ✓ Microsoft Office успешно установлен!
echo.
echo Для начала работы откройте любое приложение Office.
echo.
echo ========================================
echo.
echo [1] Вернуться в главное меню
echo [2] Выход
echo.
echo ========================================
echo.

:finish_choice
set "finish_option="
set /p "finish_option=Ваш выбор: "

if "!finish_option!"=="1" (
    goto mode_menu
) else if "!finish_option!"=="2" (
    exit /b
) else (
    cls
    echo ========================================
    echo ОШИБКА: НЕВЕРНЫЙ ВЫБОР!
    echo ========================================
    echo.
    echo Пожалуйста, введите 1 или 2
    echo.
    timeout /t 2 /nobreak >nul
    goto finish
)

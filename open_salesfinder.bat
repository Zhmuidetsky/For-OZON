@echo off
REM Запуск Python-скрипта для открытия вкладок SalesFinder
REM Проверка, что Python установлен и доступен в PATH
python --version >nul 2>&1
IF ERRORLEVEL 1 (
    echo Python не найден! Установите Python и добавьте его в PATH.
    pause
    exit /b
)

python "%~dp0open_salesfinder.py"
pause

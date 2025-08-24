@echo off
echo.
echo Начало компиляции Trubi1c.py
echo.
python.exe -m nuitka ^
    --standalone ^
    --follow-imports ^
    --enable-plugin=tk-inter ^
    --windows-console-mode=disable ^
    --lto=yes ^
    Trubi1c.py
echo.
echo Компиляция завершена
echo Папка с .exe файлом - Trubi1c.dist\
echo.
pause

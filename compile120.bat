@echo off
echo.
echo Начало компиляции Trubi120.py
echo.
python.exe -m nuitka ^
    --standalone ^
    --follow-imports ^
    --enable-plugin=tk-inter ^
    --windows-console-mode=disable ^
    --lto=yes ^
    Trubi120.py
echo.
echo Компиляция завершена
echo Папка с .exe файлом - Trubi120.dist\
echo.
pause

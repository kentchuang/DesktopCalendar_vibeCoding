@echo off
echo Building DesktopCalendar.exe...
python -m PyInstaller --onefile --noconsole --noupx --name="DesktopCalendar" main.py
if %ERRORLEVEL% NEQ 0 (
    echo Build failed! Please check the error messages above.
) else (
    echo Build successful! Check the dist\DesktopCalendar.exe file.
)
pause

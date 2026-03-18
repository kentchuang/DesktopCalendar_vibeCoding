@echo off
echo Cleaning old build files...
if exist dist\DesktopCalendar rmdir /s /q dist\DesktopCalendar
if exist build rmdir /s /q build

echo Building DesktopCalendar (onedir mode)...
python -m PyInstaller DesktopCalendar.spec
if %ERRORLEVEL% NEQ 0 (
    echo Build failed! Please check the error messages above.
) else (
    echo Build successful! Check the dist\DesktopCalendar folder.
)
pause

@echo off

echo --------
echo %company% - Fonts
echo --------

if exist "C:\Templates\%company%\Fonts" (
    xcopy "C:\Templates\%company%\Fonts\*.*" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r
)

if "%NOPAUSE%"=="" pause

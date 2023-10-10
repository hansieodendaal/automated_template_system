@echo off

echo --------
echo %company% - Letter
echo --------

xcopy "C:\Templates\%company%\%company%_Letter\%company%_Letter.dotm" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r

if "%NOPAUSE%"=="" pause

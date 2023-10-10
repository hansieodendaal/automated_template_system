@echo off

echo --------
echo %company% - Form
echo --------

xcopy "C:\Templates\%company%\%company%_Form\%company%_Form.dotm" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r

if "%NOPAUSE%"=="" pause

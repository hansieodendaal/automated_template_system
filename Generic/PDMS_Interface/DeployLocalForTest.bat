@echo off

echo ------------------
echo Generic - PDMS Interface
echo -------------------

xcopy "C:\Templates\Generic\PDMS_Interface\PDMS_Interface.xlsm" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r

if "%NOPAUSE%"=="" pause

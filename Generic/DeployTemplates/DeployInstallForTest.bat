@echo off

echo -----------
echo %company% - Install
echo -----------

xcopy "C:\Templates\%company%\Install\Install.bat" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r
xcopy "C:\Templates\%company%\Install\Fix2016UpdateError.bat" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r

if "%NOPAUSE%"=="" pause

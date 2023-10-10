@echo off

echo -----------
echo Generic - Install
echo -----------

xcopy "C:\Templates\Generic\Install\ZDoNotExecuteMe.bat" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r
xcopy "C:\Templates\Generic\Install\ReadMe.txt" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r
xcopy "C:\Templates\Generic\Install\UpdateError.vbs" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r

if "%NOPAUSE%"=="" pause

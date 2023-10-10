@echo off

echo ----------------------
echo Generic - UserManual
echo ----------------------

xcopy "C:\Templates\Generic\UserManual\UserManual_A4.pdf" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r
xcopy "C:\Templates\Generic\UserManual\UserManual_Letter.pdf" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r
xcopy "C:\Templates\Generic\UserManual\AutomatedTemplateSystemPresentation.pdf" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r


if "%NOPAUSE%"=="" pause

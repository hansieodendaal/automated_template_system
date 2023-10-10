@echo off

echo --------
echo %company% - Main
echo --------

xcopy "C:\Templates\%company%\%company%_Main\%company%_Main.dotm" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r
xcopy "C:\Templates\%company%\%company%_Main\Blank%company%doc.docx" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r

xcopy "C:\Templates\%company%\%company%_Main\Licence.txt" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r

if "%NOPAUSE%"=="" pause

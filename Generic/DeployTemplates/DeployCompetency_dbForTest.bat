@echo off

echo ------------------
echo %company% - Competency_db
echo -------------------

xcopy "C:\Templates\%company%\%company%_Competency_db\%company%_ResourcesAndCompetencies_db.xlsm" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r
xcopy "C:\Templates\%company%\%company%_Competency_db\%company%_Competency_db.xlsm" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r

if "%NOPAUSE%"=="" pause

@echo off

echo -----------------
echo %company% - Template_db
echo -----------------

xcopy "C:\Templates\%company%\%company%_Template_db\%company%_Template_db.xlsm" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r
xcopy "C:\Templates\%company%\%company%_Template_db\%company%_UserAbbreviations_db.xlsm" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r
xcopy "C:\Templates\%company%\%company%_Template_db\%company%_BusinessContacts_db.xlsm" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r
xcopy "C:\Templates\%company%\%company%_Template_db\%company%_Clauses_db.xlsm" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r

if "%NOPAUSE%"=="" pause

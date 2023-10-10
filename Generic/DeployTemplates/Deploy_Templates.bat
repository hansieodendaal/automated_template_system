@echo off

set NOPAUSE=NoPause
echo.
echo =========================================================================
echo.
echo Company = %company%
echo.

if not exist "C:\Program Files\Microsoft Office\Templates" mkdir "C:\Program Files\Microsoft Office\Templates"

if not exist "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages" mkdir "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages"

call "C:\Templates\Generic\BaseTemplate\DeployLocalForTest.bat"
echo.
call "C:\Templates\Generic\VBASecurity\DeployLocalForTest.bat"
echo.
call "C:\Templates\Generic\PDMS_Interface\DeployLocalForTest.bat"
echo.
call "C:\Templates\Generic\Install\DeployLocalForTest.bat"
echo.
call "C:\Templates\Generic\UserManual\DeployLocalForTest.bat"
echo.
call "C:\Templates\Generic\DeployTemplates\DeployInstallForTest.bat"
echo.
call "C:\Templates\Generic\DeployTemplates\DeployCompetency_dbForTest.bat"
echo.
call "C:\Templates\Generic\DeployTemplates\DeployTemplate_dbForTest.bat"
echo.
call "C:\Templates\Generic\DeployTemplates\DeployFontsForTest.bat"
echo.
if exist "C:\Templates\%company%\%company%_Main\%company%_Main.dotm" call "C:\Templates\Generic\DeployTemplates\DeployMainForTest.bat"
echo.
if exist "C:\Templates\%company%\%company%_Letter\%company%_Letter.dotm" call "C:\Templates\Generic\DeployTemplates\DeployLetterForTest.bat"
echo.
if exist "C:\Templates\%company%\%company%_Form\%company%_Form.dotm" call "C:\Templates\Generic\DeployTemplates\DeployFormForTest.bat"
echo.

echo =========================================================================
echo.
if "%NOPAUSE%"=="" pause

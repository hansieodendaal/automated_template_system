@echo off

echo Installing %Company% templates
echo.
if "%Company%" == "" goto ERROR_MSG_INSTALL: 

rem Testing installation package
echo.
echo Testing installation package
echo ----------------------------
set testFile="%installPath%ChangeWordMacroSecuritySettings.bat"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%close_word.vbs"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%close_doc.vbs"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%Reg-Office-Base.reg"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%Reg-Office-%1.reg"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%PDMS_Interface.xlsm"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%%Company%_Competency_db.xlsm"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%%Company%_ResourcesAndCompetencies_db.xlsm" 
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%%Company%_Template_db.xlsm"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%%Company%_BusinessContacts_db.xlsm"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%%Company%_Clauses_db.xlsm"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%%Company%_UserAbbreviations_db.xlsm"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%BaseTemplate.dotm"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%%Company%_*.dotm"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%Blank%Company%Doc.docx"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%UserManual*.pdf"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%AutomatedTemplateSystemPresentation.pdf"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%CopyrightNotice.txt"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%Licence.txt"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%LicenceConditions.txt"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%.\Signature\InstallSignature.txt"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%.\Signature\Signature.PNG"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%Fix2016UpdateError.bat"
  if not exist %testFile% goto ERROR_MSG_ZIP:
set testFile="%installPath%UpdateError.vbs"
  if not exist %testFile% goto ERROR_MSG_ZIP:
if exist "%installPath%ZCheckFonts.bat" (
    call "%installPath%ZCheckFonts.bat" myValue rem
    if not "!myValue!" equ "NO_ERROR" (
        echo Font files test failed: "!myValue!"
        set testFile="%installPath%ZCheckFonts.bat"
        goto ERROR_MSG_ZIP:
    )
)
echo Installation package OK
echo.

rem Registry entries for Word macro security (must be run as normal user first)
echo.
echo Registry entries for Word macro security
echo ----------------------------------------
call "%installPath%ChangeWordMacroSecuritySettings.bat" %Company%

rem Determine if running as administrator
echo.
echo Determine if running as administrator
echo -------------------------------------
set guid=%random%%random%-%random%-%random%-%random%-%random%%random%%random%
mkdir %WINDIR%\%guid%>nul 2>&1
rmdir %WINDIR%\%guid%>nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo Administrator OK
    echo.
    setlocal enableextensions
) else (
    echo Normal logged in user that is not an administrator.
	echo.
    goto INFO_MSG_ADMINISTRATOR:
)
call "%installPath%close_word.vbs"

rem Basic %Company% installation
echo.
echo Basic %Company% installation
echo ----------------------
set copyFlags= /v /y /r
rem   - Cleanup
echo   - Cleanup
echo.
set varBasic%Company%CleanupError=0
set targetFolder="C:\Program Files\Microsoft Office\Templates\%Company%"
if exist "%USERPROFILE%"\"%Company%_UserAbbreviations_db.xlsm" (
    del /f /q "%USERPROFILE%"\"%Company%_UserAbbreviations_db.xlsm"
)
if exist %targetFolder%\"%Company%_UserAbbreviations_db.xlsm" (
    xcopy %targetFolder%"\%Company%_UserAbbreviations_db.xlsm" "%USERPROFILE%" %copyFlags%
)
if exist %targetFolder% del /f /q /s %targetFolder%\*.*
if not %ERRORLEVEL% equ 0 (
    set varBasic%Company%CleanupError=1
)
echo.
rem   - Copy
echo   - Copy
echo.
set varBasic%Company%InstallationError=0
set targetFolder="C:\Program Files\Microsoft Office\Templates\%Company%"
if not exist %targetFolder% mkdir %targetFolder%
xcopy "%installPath%PDMS_Interface.xlsm" %targetFolder% %copyFlags%
xcopy "%installPath%%Company%_Competency_db.xlsm" %targetFolder% %copyFlags%
xcopy "%installPath%%Company%_ResourcesAndCompetencies_db.xlsm" %targetFolder% %copyFlags%
xcopy "%installPath%%Company%_Template_db.xlsm" %targetFolder% %copyFlags%
xcopy "%installPath%%Company%_BusinessContacts_db.xlsm" %targetFolder% %copyFlags%
xcopy "%installPath%%Company%_Clauses_db.xlsm" %targetFolder% %copyFlags%
if exist "%USERPROFILE%"\"%Company%_UserAbbreviations_db.xlsm" (
    xcopy "%USERPROFILE%"\"%Company%_UserAbbreviations_db.xlsm" %targetFolder% %copyFlags%
    del /f /q "%USERPROFILE%"\"%Company%_UserAbbreviations_db.xlsm"
    echo Re-used "%Company%_UserAbbreviations_db.xlsm"
) else (
    xcopy "%installPath%%Company%_UserAbbreviations_db.xlsm" %targetFolder% %copyFlags%
)
xcopy "%installPath%CopyrightNotice.txt" %targetFolder% %copyFlags%
xcopy "%installPath%Licence.txt" %targetFolder% %copyFlags%
xcopy "%installPath%%Company%_*.dotm" %targetFolder% %copyFlags%
xcopy "%installPath%UserManual*.pdf" %targetFolder% %copyFlags%
xcopy "%installPath%AutomatedTemplateSystemPresentation.pdf" %targetFolder% %copyFlags%
set targetFolder="C:\Program Files\Microsoft Office\Templates\%Company%\Layouts"
if not exist %targetFolder% mkdir %targetFolder%
xcopy "%installPath%Blank%Company%Doc.docx" %targetFolder% %copyFlags%
rem Install company specific fonts
if exist "%installPath%ZInstallFonts.bat" (
    echo.
    call "%installPath%ZInstallFonts.bat" NOENABLE rem no_pause_
)
rem Organizational registered layouts installation start
echo.
echo Organizational registered layouts installation start
echo ----------------------------------------------------
echo.
set CompanyLayoutsInstallPath=
xcopy "%CompanyLayoutsInstallPath%\*.doc*" %targetFolder% %copyFlags%
echo.
echo Organizational registered layouts installation stop
echo ---------------------------------------------------
rem Organizational registered layouts installation stop
if not %ERRORLEVEL% equ 0 (
    set varBasic%Company%InstallationError=1
)
echo.
rem   - Set read only attributes
echo   - Set read only attributes
echo.
set varBasic%Company%ReadOnlyAttributes=0
set targetFolder="C:\Program Files\Microsoft Office\Templates\%Company%"
attrib +R %targetFolder%\*.* /S
attrib -R %targetFolder%\"%Company%_UserAbbreviations_db.xlsm"
if not %ERRORLEVEL% equ 0 (
    set varBasic%Company%ReadOnlyAttributes=1
)
rem   - Windows XP, Windows 7, Windows 8/8.1 and Windows 10 file access rights
echo   - Windows XP, Windows 7, Windows 8/8.1 and Windows 10 file access rights
echo.
set varBasic%Company%FileAccessRightsError=0
set targetFolder="C:\Program Files\Microsoft Office\Templates\%Company%"
icacls *.>nul 2>&1
if %ERRORLEVEL% equ 0 (
    icacls %targetFolder% /t /c /grant BUILTIN\Users:m
    echo.
) else (
    cacls %targetFolder% /t /e /c /g BUILTIN\Users:c
    echo.
)
if not %ERRORLEVEL% equ 0 (
    set varBasic%Company%FileAccessRightsError=1
)

rem Base template installation
echo.
echo Base template installation
echo --------------------------
rem   - VBA project security issue start
set targetFolder="C:\Program Files\Microsoft Office\Templates\%Company%"
xcopy "%installPath%BaseTemplate.dotm" %targetFolder% %copyFlags%
rem   - VBA project security issue stop
rem   - Cleanup
echo   - Cleanup
echo.
set varBaseTemplateCleanupError=0
set targetFolder="C:\Program Files\Microsoft Office\Templates\Base"
if exist %targetFolder% del /f /q /s %targetFolder%\*.*
if not %ERRORLEVEL% equ 0 (
    set varBaseTemplateCleanupError=1
)
echo.
rem   - Copy
echo   - Copy
echo.
set varBaseTemplateInstallationError=0
set targetFolder="C:\Program Files\Microsoft Office\Templates\Base"
if not exist %targetFolder% mkdir %targetFolder%
xcopy "%installPath%CopyrightNotice.txt" %targetFolder% %copyFlags%
xcopy "%installPath%LicenceConditions.txt" %targetFolder% %copyFlags%
xcopy "%installPath%BaseTemplate.dotm" %targetFolder% %copyFlags%
xcopy "%installPath%UserManual*.pdf" %targetFolder% %copyFlags%
xcopy "%installPath%AutomatedTemplateSystemPresentation.pdf" %targetFolder% %copyFlags%
if not exist %targetFolder%\Signature mkdir %targetFolder%\Signature
xcopy "%installPath%.\Signature\Signature.PNG"  %targetFolder%\Signature %copyFlags%
xcopy "%installPath%.\Signature\InstallSignature.txt"  %targetFolder%\Signature %copyFlags%
if not %ERRORLEVEL% equ 0 (
    set varBaseTemplateInstallationError=1
)
echo.
rem   - Set read only attributes
echo   - Set read only attributes
echo.
set varBaseTemplateReadOnlyAttributes=0
set targetFolder="C:\Program Files\Microsoft Office\Templates\Base"
attrib +R %targetFolder%\*.* /S
if not %ERRORLEVEL% equ 0 (
    set varBaseTemplateReadOnlyAttributes=1
)
rem   - Windows XP, Windows 7, Windows 8/8.1 and Windows 10 file access rights
echo   - Windows XP, Windows 7, Windows 8/8.1 and Windows 10 file access rights
echo.
set varBaseTemplateFileAccessRightsError=0
set targetFolder="C:\Program Files\Microsoft Office\Templates\Base"
icacls *.>nul 2>&1
if %ERRORLEVEL% equ 0 (
    icacls %targetFolder% /t /c /grant BUILTIN\Users:m
    echo.
) else (
    cacls %targetFolder% /t /e /c /g BUILTIN\Users:c
    echo.
)
if not %ERRORLEVEL% equ 0 (
    set varBaseTemplateFileAccessRightsError=1
)

rem Set templates workgroup path
echo.
echo Set templates workgroup path
echo ----------------------------
set varSetTemplatesWorkgroupPathError=0
start /min cscript "%installPath%close_doc.vbs"
echo Opening "Blank%Company%doc.docx"
echo.
echo Please close "Blank%Company%doc.docx" if it does not close automatically.
echo.
set targetFolder="C:\Program Files\Microsoft Office\Templates\%Company%\Layouts"
%targetFolder%\"Blank%Company%doc.docx"
echo Closing "Blank%Company%doc.docx"
if not %ERRORLEVEL% equ 0 (
    set varSetTemplatesWorkgroupPathError=1
)

rem Test installation
echo.
echo Installed files
echo ---------------
rem  - Base installation
set targetFolder=C:\Program Files\Microsoft Office\Templates\Base
dir "%targetFolder%" /s/b
if not exist "%targetFolder%\BaseTemplate.dotm" goto ERROR_MSG_INSTALL_FAIL:
if not exist "%targetFolder%\CopyrightNotice.txt" goto ERROR_MSG_INSTALL_FAIL:
if not exist "%targetFolder%\LicenceConditions.txt" goto ERROR_MSG_INSTALL_FAIL:
if not exist "%targetFolder%\Signature" goto ERROR_MSG_INSTALL_FAIL:
if not exist "%targetFolder%\UserManual_*.pdf" goto ERROR_MSG_INSTALL_FAIL:
if not exist "%targetFolder%\AutomatedTemplateSystemPresentation.pdf" goto ERROR_MSG_INSTALL_FAIL:
if not exist "%targetFolder%\Signature\InstallSignature.txt" goto ERROR_MSG_INSTALL_FAIL:
if not exist "%targetFolder%\Signature\Signature.PNG" goto ERROR_MSG_INSTALL_FAIL:
rem  - Company installation
set targetFolder=C:\Program Files\Microsoft Office\Templates\%Company%
dir "%targetFolder%" /s/b
if not exist "%targetFolder%\CopyrightNotice.txt" goto ERROR_MSG_INSTALL_FAIL:
if not exist "%targetFolder%\Licence.txt" goto ERROR_MSG_INSTALL_FAIL:
if not exist "%targetFolder%\%Company%_Competency_db.xlsm" goto ERROR_MSG_INSTALL_FAIL:
if not exist "%targetFolder%\%Company%_Main.dotm" goto ERROR_MSG_INSTALL_FAIL:
if not exist "%targetFolder%\%Company%_ResourcesAndCompetencies_db.xlsm" goto ERROR_MSG_INSTALL_FAIL:
if not exist "%targetFolder%\%Company%_Template_db.xlsm" goto ERROR_MSG_INSTALL_FAIL:
if not exist "%targetFolder%\%Company%_BusinessContacts_db.xlsm" goto ERROR_MSG_INSTALL_FAIL:
if not exist "%targetFolder%\%Company%_Clauses_db.xlsm" goto ERROR_MSG_INSTALL_FAIL:
if not exist "%targetFolder%\%Company%_UserAbbreviations_db.xlsm" goto ERROR_MSG_INSTALL_FAIL:
if not exist "%targetFolder%\PDMS_Interface.xlsm" goto ERROR_MSG_INSTALL_FAIL:
if not exist "%targetFolder%\UserManual_*.pdf" goto ERROR_MSG_INSTALL_FAIL:
if not exist "%targetFolder%\AutomatedTemplateSystemPresentation.pdf" goto ERROR_MSG_INSTALL_FAIL:
set targetFolder=C:\Program Files\Microsoft Office\Templates\%Company%\Layouts
if not exist "%targetFolder%\Blank%Company%Doc.docx" goto ERROR_MSG_INSTALL_FAIL:
set "x=0"
:SymLoop1
if defined myFontFile[%x%] (
    dir "%systemroot%\fonts\!myFontFile[%x%]!" /s/b
    if not exist "%systemroot%\fonts\!myFontFile[%x%]!" (
        echo    !myFontFile[%x%]!
        goto ERROR_MSG_INSTALL_FAIL:
    )
    set /a "x+=1"
    GOTO :SymLoop1
)

rem Finish
echo.
echo.
echo Installation successful!
echo.
goto EXIT:

:INFO_MSG_ADMINISTRATOR
echo.
echo Now right click on 'Install.bat' and select 'Run as administrator'.
echo.
goto EXIT:

:ERROR_MSG_INSTALL
echo %0
echo cannot be run on its own.
echo.
echo Please run 'Install.bat' instead.
echo.
goto EXIT:

:ERROR_MSG_INSTALL_FAIL
echo.
echo.
set var
echo.
echo Installation failed!
echo Installation failed!
echo Installation failed!
echo.
echo All required files could not be installed, please contact your administrator.
echo.
goto EXIT:

:ERROR_MSG_ZIP
echo.
echo Not all installation files are present, it appears as if 
echo this install program was executed from within a zip file.
echo If that is the case, please unzip and execute from the 
echo unzipped folder/directory.
echo.
echo Command line: %0
echo Install path: %installPath%
echo Current working folder: %cd%
echo Problem file: %testFile%
echo.
goto EXIT:

:EXIT
pause
endlocal
exit
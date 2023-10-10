@echo off

if "%1" equ "" goto ERROR_MSG1:
echo.
echo Company name = '%1'
echo.

if "%2" equ "" goto ERROR_MSG2:
echo Company registry location = '%2'
echo.

if "%3" equ "" goto ERROR_MSG4:
echo Running as '%3'
echo.

rem Determine if running as administrator
if "%3" equ "Admin" (
    echo Determine if running as administrator ...
    echo.
    set guid=%random%%random%-%random%-%random%-%random%-%random%%random%%random%
    mkdir %WINDIR%\%guid%>nul 2>&1
    rmdir %WINDIR%\%guid%>nul 2>&1
    if %ERRORLEVEL% equ 0 (
        echo Administrator OK
        echo.
        setlocal enableextensions
        rem Change working folder to that of 'Install.bat'
        cd /d "%~dp0"
    ) else (
        goto ERROR_MSG3:
    )
)

call close_word.vbs

rem Microsoft Word 2013
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Word\Security /v AccessVBOM" /t REG_DWORD /d 00000001 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Word\Security /v VBAWarnings" /t REG_DWORD /d 00000001 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Word\Security\Trusted Locations\Location3" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Word\Security\Trusted Locations\Location3" /v "Path" /d "C:\\Program Files\\Microsoft Office\\Templates\\Base\\" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Word\Security\Trusted Locations\Location3" /v "AllowSubfolders" /t REG_DWORD /d 00000001 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Word\Security\Trusted Locations\Location3" /v "Date" /d "" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Word\Security\Trusted Locations\Location3" /v "Description" /d "" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Word\Security\Trusted Locations\Location%2" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Word\Security\Trusted Locations\Location%2" /v "Path" /d "C:\\Program Files\\Microsoft Office\\Templates\\%1\\" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Word\Security\Trusted Locations\Location%2" /v "AllowSubfolders" /t REG_DWORD /d 00000001 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Word\Security\Trusted Locations\Location%2" /v "Date" /d "" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Word\Security\Trusted Locations\Location%2" /v "Description" /d "" /f
echo.

rem Microsoft Word 2010
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Security" /v AccessVBOM /t REG_DWORD /d 00000001 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Security" /v VBAWarnings /t REG_DWORD /d 00000001 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Security\Trusted Locations\Location3" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Security\Trusted Locations\Location3" /v "Path" /d "C:\\Program Files\\Microsoft Office\\Templates\\Base\\" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Security\Trusted Locations\Location3" /v "AllowSubfolders" /t REG_DWORD /d 00000001 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Security\Trusted Locations\Location3" /v "Date" /d "" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Security\Trusted Locations\Location3" /v "Description" /d "" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Security\Trusted Locations\Location%2" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Security\Trusted Locations\Location%2" /v "Path" /d "C:\\Program Files\\Microsoft Office\\Templates\\%1\\" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Security\Trusted Locations\Location%2" /v "AllowSubfolders" /t REG_DWORD /d 00000001 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Security\Trusted Locations\Location%2" /v "Date" /d "" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Security\Trusted Locations\Location%2" /v "Description" /d "" /f
echo.

rem Microsoft Word 2007
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Word\Security" /v AccessVBOM /t REG_DWORD /d 00000001 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Word\Security" /v VBAWarnings /t REG_DWORD /d 00000001 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Word\Security\Trusted Locations\Location3" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Word\Security\Trusted Locations\Location3" /v "Path" /d "C:\\Program Files\\Microsoft Office\\Templates\\Base\\" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Word\Security\Trusted Locations\Location3" /v "AllowSubfolders" /t REG_DWORD /d 00000001 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Word\Security\Trusted Locations\Location3" /v "Date" /d "" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Word\Security\Trusted Locations\Location3" /v "Description" /d "" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Word\Security\Trusted Locations\Location%2" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Word\Security\Trusted Locations\Location%2" /v "Path" /d "C:\\Program Files\\Microsoft Office\\Templates\\%1\\" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Word\Security\Trusted Locations\Location%2" /v "AllowSubfolders" /t REG_DWORD /d 00000001 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Word\Security\Trusted Locations\Location%2" /v "Date" /d "" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Word\Security\Trusted Locations\Location%2" /v "Description" /d "" /f
echo.

rem Finish
echo.
goto END:

:ERROR_MSG1
echo Command line argument needed: 'Company name'
echo.
pause
goto END:

:ERROR_MSG2
echo Command line argument needed: 'Registry location (as integer .GT. 3)'
echo.
pause
goto END:

:ERROR_MSG3
echo You are not currently the administrator! (ChangeWordMacroSecuritySettings.bat)
echo.
pause
goto END:

:ERROR_MSG4
echo Command line argument needed: 'Admin / NormalUser'
echo.
pause
goto END:

:END

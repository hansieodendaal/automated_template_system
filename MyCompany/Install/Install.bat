@echo off

setlocal ENABLEDELAYEDEXPANSION

set Company=MyCompany
set installPath=%~dp0
rem Change working folder to that of 'Install.bat'
cd /d "%~dp0"

rem Determine if running as administrator
echo.
set guid=%random%%random%-%random%-%random%-%random%-%random%%random%%random%
mkdir %WINDIR%\%guid%>nul 2>&1
rmdir %WINDIR%\%guid%>nul 2>&1
if %ERRORLEVEL% equ 0 (
    setlocal enableextensions
)

if exist "%installPath%ZDoNotExecuteMe.bat" (
    call "%installPath%ZDoNotExecuteMe.bat" %Company%
) else (
    goto ERROR_MSG:
)
goto EXIT:

:ERROR_MSG
echo.
echo (%0)
echo Not all installation files are present, it appears as if 
echo this install program was executed from within a zip file.
echo If that is the case, please unzip and execute from the 
echo unzipped folder/directory.
echo.
echo Current working folder: %cd%
echo.
pause

:EXIT
exit

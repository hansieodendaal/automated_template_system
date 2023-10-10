@echo off

if not "%1" equ "NOENABLE" (
    setlocal ENABLEDELAYEDEXPANSION
    echo setlocal ENABLEDELAYEDEXPANSION, called from ZInstallFonts.bat
)

rem Verify fonts path
if not defined installPath (
    set "installPath=%~dp0"
)

rem Verify if font files exist
call "%installPath%ZCheckFonts.bat" myValue rem
if not "!myValue!" equ "NO_ERROR" (
    goto ERROR_MSG:
)

rem Determine if running as administrator
set guid=%random%%random%-%random%-%random%-%random%-%random%%random%%random%
mkdir %WINDIR%\%guid%>nul 2>&1
rmdir %WINDIR%\%guid%>nul 2>&1
if %ERRORLEVEL% equ 0 (
    set "guid=%guid%"
) else (
    goto INFO_MSG_ADMINISTRATOR:
)

rem Register fonts
set "x=0"
:SymLoop1
if defined myRegFile[%x%] (
    echo Register !myRegFile[%x%]!
    %2 echo regedit /s "%installPath%!myRegFile[%x%]!"
    regedit /s "%installPath%!myRegFile[%x%]!"

    set /a "x+=1"
    GOTO :SymLoop1
)

rem Copy fonts
set "x=0"
:SymLoop2
if defined myFontFile[%x%] (
     echo Copy !myFontFile[%x%]!
    copy /y "%installPath%!myFontFile[%x%]!" "%systemroot%\fonts"

    set /a "x+=1"
    GOTO :SymLoop2
)

rem Finish
echo Fonts installed
echo.
goto END:

:ERROR_MSG
echo.
echo All required font and/or registry files do not exist
echo !myValue!
echo.
pause
goto END:

:INFO_MSG_ADMINISTRATOR
echo.
echo Message from 'InstallFonts.bat':
echo   User '%USERNAME%' is not currently the administrator
echo   User account control (UAC) - 'Run as administrator' required
echo.
pause
goto END:

:END
if not "%3" equ "no_pause_" (
    pause
)
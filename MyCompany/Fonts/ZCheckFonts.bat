@echo off

rem Verify fonts path
if not defined installPath (
    set "installPath=%~dp0"
)

rem Get font info
call "%installPath%ZFontInfo.bat"

rem Verify if font files exist
set "FONTFILESERROR="

set "x=0"
:SymLoop1
if defined myFontFile[%x%] (
    %2 echo myFontFile[%x%] = !myFontFile[%x%]!
    %2 echo Looking for !myFontFile[%x%]!
    
    if not exist "%installPath%!myFontFile[%x%]!" (
        set "FONTFILESERROR=%FONTFILESERROR% !myFontFile[%x%]!"
        echo Could not find !myFontFile[%x%]!
    ) else (
        %2 echo Found !myFontFile[%x%]!
    )
    
    set /a "x+=1"
    GOTO :SymLoop1
)

set "x=0"
:SymLoop2
if defined myRegFile[%x%] (
    %2 echo myRegFile[%x%] = !myRegFile[%x%]!
    %2 echo Looking for !myRegFile[%x%]!
    
    if not exist "%installPath%!myRegFile[%x%]!" (
        set "FONTFILESERROR=%FONTFILESERROR% !myRegFile[%x%]!"
        echo Could not find !myRegFile[%x%]!
    ) else (
        %2 echo Found !myRegFile[%x%]!
    )
    
    set /a "x+=1"
    GOTO :SymLoop2
)

if defined FONTFILESERROR (
    set "FONTFILESERROR=ERROR file not found; %FONTFILESERROR%"
) else (
    set "FONTFILESERROR=NO_ERROR"
)
set "%~1=%FONTFILESERROR%"


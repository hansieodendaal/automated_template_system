@echo off

if "%1" equ "" goto ERROR_MSG1:
echo Company name = '%1'
echo.

call "%installPath%close_word.vbs"

rem Import Microsoft Word 2007, 2010, 2013, 2016 registry settings
if exist "%installPath%Reg-Office-%1.reg" (
    echo Microsoft Word 2007, 2010, 2013, 2016: Reg-Office-%1.reg
    regedit /s "%installPath%Reg-Office-%1.reg"
    echo.
) else (
    goto ERROR_MSG2:
)
if exist "%installPath%Reg-Office-Base.reg" (
    echo Microsoft Word 2007, 2010, 2013, 2016: Reg-Office-Base.reg
    regedit /s "%installPath%Reg-Office-Base.reg"
    echo.
) else (
    goto ERROR_MSG4:
)

rem Finish
echo User settings updated.
echo.
goto END:

:ERROR_MSG1
echo.
echo Command line argument needed: 'Company name'
echo.
pause
goto END:

:ERROR_MSG2
echo.
echo Registry file does not exist: 'Reg-Office-%1.reg'
echo.
pause
goto END:

:ERROR_MSG4
echo.
echo Registry file does not exist: 'Reg-Office-Base.reg'
echo.
pause
goto END:

:ERROR_MSG3
echo.
echo You are not currently the administrator! (ChangeWordMacroSecuritySettings.bat)
echo.
pause
goto END:

:END

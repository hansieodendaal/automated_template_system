@echo off
REM ********************************************************************
REM  This batch file starts a command shell under the current user account,
REM  after temporarily adding that user to the local Administrators group.
REM  Any program launched from that command shell will also run with
REM  administrative privileges.
REM
REM  You will be prompted for two passwords in two separate command shells:
REM  first, for the password of the local administrator account, and
REM  second for the password of the account under which you are logged on.
REM  (The reason for this is that you are creating a new logon session in
REM  which the user will be a member of the Administrators group.)
REM
REM  CUSTOMIZATION:
REM  The following values may be changed in order to customize this script:
REM
REM  * _Prog_  : the program to run
REM
REM  * _Admin_ : the name of the administrative account that can make changes
REM              to local groups (usu. "Administrator" unless you renamed the
REM              local administrator account).  The first password prompt
REM              will be for this account.
REM
REM  * _Group_ : the local group to temporarily add the user to (e.g.,
REM              "Administrators").
REM
REM  * _User_  : the account under which to run the new program.  The second
REM              password prompt will be for this account.  Leave it as
REM              %USERDOMAIN%\%USERNAME% in order to elevate the current user.
REM ********************************************************************

setlocal
rem set _Admin_=%COMPUTERNAME%\Administrator
set _Admin_=%COMPUTERNAME%\AdminOnHansiePC
set _Group_=Administrators
set _Prog_="cmd.exe /k Title *** %* as Admin *** && cd c:\ && color 4F"
set _User_=%USERDOMAIN%\%USERNAME%

if "%1"=="" (
	runas /u:%_Admin_% "%~s0 %_User_%"
	if ERRORLEVEL 1 echo. && pause
) else (
	echo Adding user %* to group %_Group_%...
	net localgroup %_Group_% "%*" /ADD
	if ERRORLEVEL 1 echo. && pause
	echo.
	echo Starting program in new logon session...
	runas /u:"%*" %_Prog_%
	if ERRORLEVEL 1 echo. && pause
	echo.
	echo Removing user %* from group %_Group_%...
	net localgroup %_Group_% "%*" /DELETE
	if ERRORLEVEL 1 echo. && pause
)
endlocal

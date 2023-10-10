@echo off

rem Clear variables
set "x=0"
:SymLoop1
if defined myFontFile[%x%] (
    set myFontFile[%x%]=
    set /a "x+=1"
    GOTO :SymLoop1
)

set "x=0"
:SymLoop2
if defined myRegFile[%x%] (
    set myRegFile[%x%]=
    set /a "x+=1"
    GOTO :SymLoop2
)
set myFile=

rem Setup font file array
set Font=Rondo
set myFontFile[0]=%Font%.TTF
set myRegFile[0]=%Font%.reg

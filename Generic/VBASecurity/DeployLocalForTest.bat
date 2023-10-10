@echo ---------------------
@echo Generic - VBASecurity
@echo ---------------------

xcopy "C:\Templates\Generic\VBASecurity\ChangeWordMacroSecuritySettings.bat" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r
xcopy "C:\Templates\Generic\VBASecurity\Reg-Office-Base.reg" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r
xcopy "C:\Templates\Generic\VBASecurity\Reg-Office-%company%.reg" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r
xcopy "C:\Templates\Generic\VBASecurity\close_word.vbs" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r
xcopy "C:\Templates\Generic\VBASecurity\close_doc.vbs" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r

@if "%NOPAUSE%"=="" pause

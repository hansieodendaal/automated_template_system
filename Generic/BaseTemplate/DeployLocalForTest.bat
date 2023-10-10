@echo ----------------------
@echo Generic - BaseTemplate
@echo ----------------------

xcopy "C:\Templates\Generic\BaseTemplate\BaseTemplate.dotm" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r
xcopy "C:\Templates\Generic\BaseTemplate\InstallSignature.txt" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\Signature\*.*" /v /y /r
xcopy "C:\Templates\Generic\BaseTemplate\Signature.PNG" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\Signature\*.*" /v /y /r
xcopy "C:\Templates\Generic\BaseTemplate\CopyrightNotice.txt" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r
xcopy "C:\Templates\Generic\BaseTemplate\LicenceConditions.txt" "C:\Program Files\Microsoft Office\Templates\zzInstallationPackages\%company%_Templates_Install\*.*" /v /y /r

@if "%NOPAUSE%"=="" pause

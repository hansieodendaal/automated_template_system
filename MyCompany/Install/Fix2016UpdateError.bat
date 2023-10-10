@echo off

set Company=MyCompany

set targetFolder=C:\Program Files\Microsoft Office\Templates\%Company%\Layouts
set targetDocument="%targetFolder%\Blank%Company%doc.docx"

start UpdateError.vbs

ping localhost -n 3 > nul

%targetDocument%

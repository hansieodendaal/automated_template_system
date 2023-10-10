@echo off

cls

if exist "C:\Program Files (x86)\ImageMagick-6.8.8-Q8\convert.exe" set IM_convert="C:\Program Files (x86)\ImageMagick-6.8.8-Q8\convert.exe"
if exist "C:\Program Files\ImageMagick-6.8.8-Q8\convert.exe" set IM_convert="C:\Program Files\ImageMagick-6.8.8-Q8\convert.exe"
set IM_convert

if exist "C:\Program Files (x86)\ImageMagick-6.8.8-Q8\composite.exe" set IM_composite="C:\Program Files (x86)\ImageMagick-6.8.8-Q8\composite.exe"
if exist "C:\Program Files\ImageMagick-6.8.8-Q8\composite.exe" set IM_composite="C:\Program Files\ImageMagick-6.8.8-Q8\composite.exe"
set IM_composite

%IM_convert% Signature.PNG -stroke "#00AA0090" -strokewidth 10 -draw "line 70,116 121,213" -draw "line 121,213 370,7" -draw "line 370,7 140,155" -stroke green -draw "line 140,155 70,116" -stroke black -pointsize 23 -strokewidth 0 -draw "text 62,182 'Hansie Odendaal'" -draw "text 62,212 '2014-01-17, 7:06 AM'" CompositeSignature.PNG 


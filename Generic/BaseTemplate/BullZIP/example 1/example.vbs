Rem -- This example will show you how to create a very simple runonce configuration.
Option Explicit

Dim printername
Dim output, statusfile, fso, currentdir, documentfile, util, settings

Rem -- Create the COM helper object.
set util = CreateObject("biopdf.PdfUtil")

Rem -- Create the COM settings object to control the printer.
set settings = CreateObject("biopdf.PdfSettings")

Rem -- Get current path of this script.
Set fso = CreateObject("Scripting.FileSystemObject")
currentdir = fso.GetAbsolutePathName(".")

output = currentdir & "\out\example.pdf"
statusfile = currentdir & "\out\status.ini"
documentfile = currentdir & "\in\example.rtf"

Rem -- Change the value of printer name if you want to use another PDF printer
printername = util.DefaultPrinterName

settings.PrinterName = printername
settings.SetValue "Output", output
settings.SetValue "WatermarkText", Now
settings.SetValue "WatermarkColor", "#FF9900"
settings.SetValue "ShowSettings", "never"
settings.SetValue "ShowPDF", "no"
settings.SetValue "ShowProgress", "no"
settings.SetValue "ShowProgressFinished", "no"
settings.SetValue "ConfirmOverwrite", "no"
settings.SetValue "StatusFile", statusfile
settings.SetValue "StatusFileEncoding", "unicode"

Rem -- Write settings to the runonce.ini.
settings.WriteSettings True

Rem -- Remove old output and status files
if fso.FileExists(output) then fso.DeleteFile(output)
if fso.FileExists(statusfile) then fso.DeleteFile(statusfile)

Rem -- Print the file
util.PrintFile documentfile, printername

Rem -- Wait for the status file. 
Rem -- The printing has finished when the status file is written.
if util.WaitForFile(statusfile, 10000) then
	wscript.echo "Print job finished."
else
	wscript.echo "Print job timed out."
end if

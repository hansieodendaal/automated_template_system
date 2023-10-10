Rem -- This example will show you how to create a very simple runonce configuration.

Rem -- Get current path of this script.
Set fso = CreateObject("Scripting.FileSystemObject")
currentdir = fso.GetAbsolutePathName(".")

Rem -- Read the info xml
Set xmldom = CreateObject("MSXML.DOMDocument")
xmldom.Load(currentdir & "\info.xml")

Rem -- Get the program id of the automation object.
progid = xmldom.SelectSingleNode("/xml/progid").text

Rem -- Create the COM object to control the printer.
set obj = CreateObject(progid)

Rem -- Get the default printer name.
Rem -- You can override this setting to specify a specific printer.
printername = obj.GetPrinterName

runonce = obj.GetSettingsFileName(true)

Rem -- Print all the files in the 'in' folder
Set fldr = fso.GetFolder(currentdir & "\in")
cnt = 0
For Each f In fldr.files
	cnt = cnt + 1
	output = currentdir & "\out\" & Replace(f.name, ".txt", "") & ".pdf"

	Rem -- Set the values
	obj.Init
	obj.SetValue "Output", output
	obj.SetValue "ShowSettings", "never"
	obj.SetValue "ShowPDF", "no"
	obj.SetValue "WatermarkText", now
	obj.SetValue "ShowProgress", "no"
	obj.SetValue "ShowProgressFinished", "no"
	obj.SetValue "SuppressErrors", "yes"
	obj.SetValue "ConfirmOverwrite", "no"

	Rem -- Write settings to the runonce-Invoice.ini
	obj.WriteSettings True

	Rem -- Print the document
	printfile = currentdir & "\in\" & f.name
	cmd = """" & currentdir & "\printto.exe"" """ & printfile & """ """ & printername & """"
	
	Set WshShell = WScript.CreateObject("WScript.Shell")
	ret = WshShell.Run(cmd, 1, true)

	Rem -- Wait until the runonce is removed.
	Rem -- When the runonce is removed it means that the gui.exe program 
	rem -- has picked up the print job and is ready for the next.
	While fso.fileexists(runonce)
		Rem -- Wait for some milliseconds before testing again
		wscript.sleep 100
	Wend
Next

Rem -- Dispose the printer control object
set obj = Nothing

Wscript.Echo cnt & " documents were printed."
WorkingDir = "Z:\QA-SHARE\CORE II Plan Templates\Weekly Send to HealthX"

extension = ".XLSX"
extension2 = ".CSV"

Dim fso, myFolder, fileColl, aFile, FileName, SaveName, objShell
Dim objExcel, objWorkbook

set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FilesystemObject")
Set myFolder = fso.GetFolder(WorkingDir)
Set fileColl = myFolder.Files

Set objExcel = CreateObject("Excel.Application")

Set lfso = WScript.CreateObject("Scripting.Filesystemobject")
Set l = lfso.OpenTextFile("Z:\Project Plans\HealthX\CORE II\Plan Templates BATCHQ\Templates\excel to csv log.txt", 2)
Set l2 = lfso.OpenTextFile("Z:\Project Plans\HealthX\CORE II\Plan Templates BATCHQ\Templates\remove lines log.txt", 2)

objExcel.Visible = False
objExcel.DisplayAlerts = False

For Each aFile In fileColl
    ext = Right(aFile.Name, 5)
        If UCase(ext) = UCase(extension) Then
            FileName = Left(aFile,InStrRev(aFile,"."))
            Set objWorkbook = objExcel.Workbooks.Open(aFile)
            SaveName = FileName & "csv"
			l.WriteLine Now & " " & aFile & " changed to .csv"
            objWorkbook.SaveAs SaveName, 23
            objWorkbook.Close 
		Elseif UCase(ext) <> UCase(extension) Then
			On Error Resume Next
			l.WriteLine Now & " " & aFile & " IS NOT A .XLS!!"
		End If	
Next

For Each aFile In fileColl
    ext = Right(aFile.Name ,4)
        If UCase(ext) = UCase(extension2) Then
			Set objWorkbook = objExcel.Workbooks.Open(aFile)
			i = 1
			Do Until objExcel.Cells(i, 2).Value = "" 
				If objExcel.Cells(i, 2).Value = "PlanID" Then
				Set objRange = objExcel.Cells(i, 2).EntireRow
				objRange.Delete                           'deletes header row
				l2.WriteLine Now & " " & aFile & " updated."
				End If
			i = i + 1
			Loop
			objWorkbook.Save
			objWorkbook.Close			
		Elseif UCase(ext) <> UCase(extension2) Then
			On Error Resume Next
		End If	
Next

l.WriteLine Now & " All Done!"
l.Close
Wscript.Echo "Done!"
l2.WriteLine Now & " All Done!"
l2.Close
Wscript.Echo "All Done!"
Set objWorkbook = Nothing
Set objExcel = Nothing
Set fso = Nothing
Set myFolder = Nothing
Set fileColl = Nothing

objShell.Run "%comspec% /k Z:",,false
WScript.Sleep 100

objShell.SendKeys "cd QA-SHARE\CORE II Plan Templates\Weekly Send to HealthX{ENTER}"
Wscript.Sleep 100

objShell.SendKeys "copy *.csv newfile.csv{ENTER}"
Wscript.Sleep 1000

objShell.SendKeys "exit{ENTER}"

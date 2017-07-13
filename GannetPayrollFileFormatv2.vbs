Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Dim sArray

Set objFSO = CreateObject("Scripting.FileSystemObject")

Function myDateFormat(myDate)
    d = WhatEver(Day(myDate))
    m = WhatEver(Month(myDate))    
    y = Year(myDate)
    myDateFormat= y & m & d
End Function

Function WhatEver(num)
    If(Len(num) = 1) Then
        WhatEver = "0" & num
    Else
        WhatEver = num
    End If
End Function


sSourceFile = InputBox("Enter in the full filename, include the path")


Set objTextFile = objFSO.OpenTextFile(sSourceFile, ForReading)

If Not objFSO.FolderExists("Z:\billing\CEW\NEHPGAN\" & Year(Now)) Then
	MsgBox "Z:\billing\CEW\NEHPGAN\" & Year(Now) & " folder does not exist. Creating now."
	objFSO.CreateFolder("Z:\billing\CEW\NEHPGAN\" & Year(Now))
End If


Set sOutputFile = objFSO.CreateTextFile("Z:\billing\CEW\NEHPGAN\" & Year(Now) & "\NEHPGANPayrollFile_" & myDateFormat(Now) & ".txt", ForWriting)


sOutPutFile.Writeline "PLAN ID" & "," & "SSN" & "," & "Participant ID" & "," & "Payroll Clock" & "," & "DEN DEDCODE" & "," & "DEN" & "," & "MED DED CODE" & "," & "MED" & "," & "VIS DEDCODE" & "," & "VIS" & _
"," & "Division"

Do Until objTextFile.AtEndOfStream
	sRecord = objTextFile.Readline
	sSSN = ""
	sArray = Split(sRecord, ",")
	If sArray(19) = """""" Then
		sDenDed = """0"""
		Else
		sDenDed = sArray(19)
	End If
	If sArray(13) = """""" Then
		sMedDed = """0"""
		Else
		sMedDed = sArray(13)
	End If
	If sArray(25) = """""" Then
		sVisDed = """0"""
		Else
		sVisDed = sArray(25)
	End If
	If Len(sArray(5)) < 9 Then
		sSSN = "0" & sArray(5)
		Else
		sSSN = sArray(5)
	End If
	sOutputFile.Writeline "HW" & "," & sSSN & "," & sArray(1) & "," & sArray(7) & "," & "02" & "," & sDenDed & "," & "01" & "," & sMedDed & "," & "15" & "," & sVisDed & "," & sArray(6)
Loop

sOutputFile.Close
objTextFile.Close

MsgBox "Done"

Set sOutputFile = Nothing
Set objTextFile = Nothing
Set objFSO = Nothing

	
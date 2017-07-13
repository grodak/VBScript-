Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=C:\Users\grodak\Documents\ODBC.mdb"

Set objConn = CreateObject("ADODB.Connection")
Set objFSO = CreateObject("Scripting.FileSystemObject")

sGroupID = InputBox("Enter LuminX GroupID:", "CEEM Participant & Coverage Extract")
If sGroupID = "" Then
	MsgBox "No Group was entered. Script will exit.", vbCritical
	WScript.Quit
End If

If Not objFSO.FolderExists("H:\CEEM\" & sGroupID) Then
	objFSO.CreateFolder("H:\CEEM\" & sGroupID)
End If

objConn.Open connStr

StartTime = Timer()

Set objEligTextFile = objFSO.CreateTextFile("H:\CEEM\" & sGroupID & "\" & sGroupID & "Part4.txt", ForWriting)
Set objCovTextFile = objFSO.CreateTextFIle("H:\CEEM\" & sGroupID & "\" & sGroupID & "PartCov4.txt", ForWriting)

'====================================
'Writes Header for Eligibility file. 
'====================================
objEligTextFile.Writeline "Id" & vbTab & "Group_Number" & vbTab & "Effective_From" & vbTab & "Effective_To" & vbTab & "Benefits_Effective" & vbTab & "Benefits_End" & vbTab & "Division" & vbTab & _
"First_Name" & vbTab & "MI" & vbTab & "Last_Name" & vbTab & "Date_of_Birth" & vbTab & "SSN" & vbTab & "Gender" & vbTab & "Marital_Status" & vbTab & "Smoker" & vbTab & "Disabled" & vbTab & _
"Date_of_Hire" & vbTab & "Status" & vbTab & "Address" & vbTab & "Address2" & vbTab & "City" & vbTab & "State" & vbTab & "Zip" & vbTab & "Annual_Salary" & vbTab & "Weekly_Hours" & vbTab & _
"Class" & vbTab & "Occupation" & vbTab & "Location"

'================================
'Writes Header for Coverage file. 
'================================
objCovTextFile.Writeline "Group_Number" & vbTab & "Division" & vbTab & "Effective_From" & vbTab & "Effective_To" & vbTab & "Participant" & vbTab & "Product" & vbTab & _
"Plan" & vbTab & "Enrollment" & vbTab & "Volume"
	
	'eSQL = "SELECT GPPARTIC, GPGROUP, GPINITDT, GPBENEND, GPBENEFF, GPBENEND, GPDIV, GPFIRST, GPMI, GPLAST, GPDOB, GPSSN, GPSEX, GPMARSTA, GPSMOKE, GPDISABL, GPHIREDT, " & _
	'"IIf(Len(Trim(GPBENEND))>0,'T','A') AS Status, GPSTNUMB & GPSTDIRC & GPSTNAME & IIf(InStr(GPSTSUFF,'%'),Left(GPSTSUFF,InStr(GPSTSUFF,'%')-1),GPSTSUFF) AS Address, " & _
	'"GPADDR2, GPCITY, GPSTATE, GPZIP, GPCLOCK, G7USEDF6, G7PAYID, G7PAYCYC, GPSALARY, G7HRWEEK FROM PARTICIP INNER JOIN PARTMEDC ON PARTICIP.GPGROUP = PARTMEDC.G7GROUP " & _
	'"AND PARTICIP.GPPARTIC = PARTMEDC.G7PARTIC WHERE GPGROUP = '" & sGroupID & "' AND G7GROUP ='" & sGroupID & "'"
	
	'This is the "standard" eligibility query. 
	eSQL = "SELECT GPPARTIC, GPGROUP, GPINITDT, GPBENEND, GPBENEFF, GPBENEND, GPDIV, GPFIRST, GPMI, GPLAST, GPDOB, GPSSN, GPSEX, GPMARSTA, GPSMOKE, GPDISABL, GPHIREDT, " & _
	"IIf(Len(Trim(GPBENEND))>0,'T','A') AS Status, GPSTNUMB & GPSTDIRC & GPSTNAME & IIf(InStr(GPSTSUFF,'%'),Left(GPSTSUFF,InStr(GPSTSUFF,'%')-1),GPSTSUFF) AS Address, " & _
	"GPADDR2, GPCITY, GPSTATE, GPZIP, GPSALARY, G7HRWEEK FROM PARTICIP INNER JOIN PARTMEDC ON PARTICIP.GPGROUP = PARTMEDC.G7GROUP " & _
	"AND PARTICIP.GPPARTIC = PARTMEDC.G7PARTIC WHERE PARTICIP.GPGROUP = '" & sGroupID & "' AND PARTMEDC.G7GROUP ='" & sGroupID & "'"

	'This is the extract for NEHPGAN
	'eSQL = "SELECT PARTICIP.GPPARTIC, PARTICIP.GPGROUP, PARTICIP.GPINITDT, PARTICIP.GPBENEND, PARTICIP.GPBENEFF, PARTICIP.GPBENEND, PARTICIP.GPDIV, " & _
	'"PARTICIP.GPFIRST, PARTICIP.GPMI, PARTICIP.GPLAST, PARTICIP.GPDOB, PARTICIP.GPSSN, PARTICIP.GPSEX, PARTICIP.GPMARSTA, PARTICIP.GPSMOKE, PARTICIP.GPDISABL, PARTICIP.GPHIREDT, " & _
	'"IIf(Len(Trim(PARTICIP.GPBENEND))>0,'T','A') AS Status, PARTICIP.GPSTNUMB & PARTICIP.GPSTDIRC & PARTICIP.GPSTNAME & " & _
	'"IIf(InStr(PARTICIP.GPSTSUFF,'%'),Left(PARTICIP.GPSTSUFF,InStr(PARTICIP.GPSTSUFF,'%')-1),PARTICIP.GPSTSUFF) AS Address, " & _
	'"PARTICIP.GPADDR2, PARTICIP.GPCITY, PARTICIP.GPSTATE, PARTICIP.GPZIP, PARTICIP.GPSALARY, PARTMEDC.G7HRWEEK, IIf(PARTMEDC.G7USEDF6 = 'Y','2','1') As Class, " & _
	'"IIf(PARTMEDC.G7PAYID = 'D','Unpaid Leave','Active') As Occupation, Trim(GPCLOCK) & 'A' As Location " & _
	'"FROM PARTICIP INNER JOIN PARTMEDC ON PARTICIP.GPGROUP = PARTMEDC.G7GROUP AND PARTICIP.GPPARTIC = PARTMEDC.G7PARTIC WHERE " & _
	'"PARTICIP.GPGROUP = '" & sGroupID & "' AND PARTMEDC.G7GROUP ='" & sGroupID & "'"
Set rsElig = objConn.execute(eSQL)
counter = 0
Do While Not rsElig.EOF
	counter = counter + 1
	'==========================================================================================================================
	'= Grabs ID and checks against last ID record. If new, writes the record. This is to grab the last record in the dataset. =
	'==========================================================================================================================
	If rsElig.Fields(0) = lstID Then
		lstID = rsElig.Fields(0)
		rsElig.MoveNext
		Else
				'===========================================================
				'Grabs each ID from eSQL record set and finds coverage info. 
				'LIF (LI) = rsCov.Fields(54)                                 
				'SPL (SL) =                                                  
				'DPL (DL) =                                                  
				'ADD (AD) = rsCov.Fields(55)                                 
				'STD (ST) = rsCov.Fields(56)                                 
				'LTD (LT) = rsCov.Fields(53)                                 
				'SUP (M1) = rsCov.Fields(57)                                 
				'SSP (M3) = rsCov.Fields(58)                                 
				'SDP (M4) = rsCov.Fields(59)                                 
				'SAD (M2) = rsCov.Fields(60)                                 
				'SM2 (M5) = rsCov.Fields(61)                                 
				'DM2 (M6) = rsCov.Fields(62)                                 
				'===========================================================
				strPartID = rsElig.Fields(0)
				cSQL = "SELECT PEGROUP, PEDIV, PEFROMDT, PETODATE, PEPARTIC, PECOVCAT1, PECOVCAT2, PECOVCAT3, PECOVCAT4, PECOVCAT5, PECOVCAT6, PECOVCAT7, PECOVCAT8, " & _
				"PECOVCAT9, PECOVCAT10, PECOVCAT11, PECOVCAT12, PECOVCAT13, PECOVCAT14, PECOVCAT15, PECOVCAT16, PEPLAN1, PEPLAN2, PEPLAN3, PEPLAN4, PEPLAN5, PEPLAN6, " & _
				"PEPLAN7, PEPLAN8, PEPLAN9, PEPLAN10, PEPLAN11, PEPLAN12, PEPLAN13, PEPLAN14, PEPLAN15, PEPLAN16, PEENRLEV1, PEENRLEV2, " & _
				"PEENRLEV3, PEENRLEV4, PEENRLEV5, PEENRLEV6, PEENRLEV7, PEENRLEV8, PEENRLEV9, PEENRLEV10, PEENRLEV11, PEENRLEV12, PEENRLEV13, PEENRLEV14, PEENRLEV15, PEENRLEV16, GPVELTD, " & _
				"GPVELIFE, GPVEADD, GPVESTD, GPVEMI1, GPVSMI1, GPVDMI1, GPVEMI2, GPVSMI2, GPVDMI2 " & _
				"FROM PARTCOVG INNER JOIN PARTICIP ON PARTCOVG.PEGROUP = PARTICIP.GPGROUP AND PARTCOVG.PEPARTIC = PARTICIP.GPPARTIC WHERE PEGROUP = '" & sGroupID & _
				"' AND PEPARTIC = '" & strPartID & "' ORDER BY PEPARTIC, PETODATE DESC;"
				
				Set rsCov = objConn.execute(cSQL)
				Do While Not rsCov.EOF
				'======================================================================================================================
				'Grabs ID and checks against last ID record. If new, writes the record. This is to grab the last record in the dataset.
				'======================================================================================================================
				If rsCov.Fields(4) = lstID Then
					lstID = rsCov.Fields(4)
					rsCov.MoveNext
						Else
							'===============================================================================================
							'Checks each PECOVCAT field to see if blank. If not, writes the Cov, Plan, and Enrollment Level. 
							'This also checks for volume based benefits, and if there is a match, it pulls the corresponding 
							'volume. **THIS IS SET FOR THE ABS EMPLOYEES GROUP. IT NEEDS TO BE VALIDATED FOR OTHER GROUPS**  
							'===============================================================================================
							sID = rsCov.Fields(0) & vbTab & rsCov.Fields(1) & vbTab & rsCov.Fields(2) & vbTab & rsCov.Fields(3) & vbTab & rsCov.Fields(4)
							If Len(Trim(rsCov.Fields(5))) > 0 Then
								Select Case rsCov.Fields(5)
								Case "ADD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(5) & vbTab & rsCov.Fields(21) & vbTab & rsCov.Fields(37) & vbTab & rsCov.Fields(55)
								Case "LIF"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(5) & vbTab & rsCov.Fields(21) & vbTab & rsCov.Fields(37) & vbTab & rsCov.Fields(54)
								Case "LTD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(5) & vbTab & rsCov.Fields(21) & vbTab & rsCov.Fields(37) & vbTab & rsCov.Fields(53)
								Case "STD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(5) & vbTab & rsCov.Fields(21) & vbTab & rsCov.Fields(37) & vbTab & rsCov.Fields(56)
								Case "SUP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(5) & vbTab & rsCov.Fields(21) & vbTab & rsCov.Fields(37) & vbTab & rsCov.Fields(57)
								Case "SSP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(5) & vbTab & rsCov.Fields(21) & vbTab & rsCov.Fields(37) & vbTab & rsCov.Fields(58)
								Case "SDP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(5) & vbTab & rsCov.Fields(21) & vbTab & rsCov.Fields(37) & vbTab & rsCov.Fields(59)
								Case "SAD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(5) & vbTab & rsCov.Fields(21) & vbTab & rsCov.Fields(37) & vbTab & rsCov.Fields(60)
								Case "SM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(5) & vbTab & rsCov.Fields(21) & vbTab & rsCov.Fields(37) & vbTab & rsCov.Fields(61)
								Case "DM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(5) & vbTab & rsCov.Fields(21) & vbTab & rsCov.Fields(37) & vbTab & rsCov.Fields(62)								
								Case Else
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(5) & vbTab & rsCov.Fields(21) & vbTab & rsCov.Fields(37)
								End Select
							End If
							If Len(Trim(rsCov.Fields(6))) > 0 Then 
								Select Case rsCov.Fields(6)
								Case "ADD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(6) & vbTab & rsCov.Fields(22) & vbTab & rsCov.Fields(38) & vbTab & rsCov.Fields(55)
								Case "LIF"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(6) & vbTab & rsCov.Fields(22) & vbTab & rsCov.Fields(38) & vbTab & rsCov.Fields(54)
								Case "LTD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(6) & vbTab & rsCov.Fields(22) & vbTab & rsCov.Fields(38) & vbTab & rsCov.Fields(53)
								Case "STD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(6) & vbTab & rsCov.Fields(22) & vbTab & rsCov.Fields(38) & vbTab & rsCov.Fields(56)
								Case "SUP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(6) & vbTab & rsCov.Fields(22) & vbTab & rsCov.Fields(38) & vbTab & rsCov.Fields(57)
								Case "SSP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(6) & vbTab & rsCov.Fields(22) & vbTab & rsCov.Fields(38) & vbTab & rsCov.Fields(58)
								Case "SDP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(6) & vbTab & rsCov.Fields(22) & vbTab & rsCov.Fields(38) & vbTab & rsCov.Fields(59)
								Case "SAD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(6) & vbTab & rsCov.Fields(22) & vbTab & rsCov.Fields(38) & vbTab & rsCov.Fields(60)
								Case "SM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(6) & vbTab & rsCov.Fields(22) & vbTab & rsCov.Fields(38) & vbTab & rsCov.Fields(61)
								Case "DM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(6) & vbTab & rsCov.Fields(22) & vbTab & rsCov.Fields(38) & vbTab & rsCov.Fields(62)								
								Case Else
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(6) & vbTab & rsCov.Fields(22) & vbTab & rsCov.Fields(38)
								End Select
							End If
							If Len(Trim(rsCov.Fields(7))) > 0 Then
								Select Case rsCov.Fields(7)
								Case "ADD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(7) & vbTab & rsCov.Fields(23) & vbTab & rsCov.Fields(39) & vbTab & rsCov.Fields(55)
								Case "LIF"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(7) & vbTab & rsCov.Fields(23) & vbTab & rsCov.Fields(39) & vbTab & rsCov.Fields(54)
								Case "LTD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(7) & vbTab & rsCov.Fields(23) & vbTab & rsCov.Fields(39) & vbTab & rsCov.Fields(53)
								Case "STD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(7) & vbTab & rsCov.Fields(23) & vbTab & rsCov.Fields(39) & vbTab & rsCov.Fields(56)
								Case "SUP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(7) & vbTab & rsCov.Fields(23) & vbTab & rsCov.Fields(39) & vbTab & rsCov.Fields(57)
								Case "SSP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(7) & vbTab & rsCov.Fields(23) & vbTab & rsCov.Fields(39) & vbTab & rsCov.Fields(58)
								Case "SDP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(7) & vbTab & rsCov.Fields(23) & vbTab & rsCov.Fields(39) & vbTab & rsCov.Fields(59)
								Case "SAD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(7) & vbTab & rsCov.Fields(23) & vbTab & rsCov.Fields(39) & vbTab & rsCov.Fields(60)
								Case "SM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(7) & vbTab & rsCov.Fields(23) & vbTab & rsCov.Fields(39) & vbTab & rsCov.Fields(61)
								Case "DM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(7) & vbTab & rsCov.Fields(23) & vbTab & rsCov.Fields(39) & vbTab & rsCov.Fields(62)
								Case Else
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(7) & vbTab & rsCov.Fields(23) & vbTab & rsCov.Fields(39)
								End Select
							End If
							If Len(Trim(rsCov.Fields(8))) > 0 Then
								Select Case rsCov.Fields(8)
								Case "ADD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(8) & vbTab & rsCov.Fields(24) & vbTab & rsCov.Fields(40) & vbTab & rsCov.Fields(55)
								Case "LIF"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(8) & vbTab & rsCov.Fields(24) & vbTab & rsCov.Fields(40) & vbTab & rsCov.Fields(54)
								Case "LTD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(8) & vbTab & rsCov.Fields(24) & vbTab & rsCov.Fields(40) & vbTab & rsCov.Fields(53)
								Case "STD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(8) & vbTab & rsCov.Fields(24) & vbTab & rsCov.Fields(40) & vbTab & rsCov.Fields(56)
								Case "SUP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(8) & vbTab & rsCov.Fields(24) & vbTab & rsCov.Fields(40) & vbTab & rsCov.Fields(57)
								Case "SSP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(8) & vbTab & rsCov.Fields(24) & vbTab & rsCov.Fields(40) & vbTab & rsCov.Fields(58)
								Case "SDP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(8) & vbTab & rsCov.Fields(24) & vbTab & rsCov.Fields(40) & vbTab & rsCov.Fields(59)
								Case "SAD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(8) & vbTab & rsCov.Fields(24) & vbTab & rsCov.Fields(40) & vbTab & rsCov.Fields(60)
								Case "SM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(8) & vbTab & rsCov.Fields(24) & vbTab & rsCov.Fields(40) & vbTab & rsCov.Fields(61)
								Case "DM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(8) & vbTab & rsCov.Fields(24) & vbTab & rsCov.Fields(40) & vbTab & rsCov.Fields(62)
								Case Else
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(8) & vbTab & rsCov.Fields(24) & vbTab & rsCov.Fields(40)
								End Select
							End If
							If Len(Trim(rsCov.Fields(9))) > 0 Then
								Select Case rsCov.Fields(9)
								Case "ADD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(9) & vbTab & rsCov.Fields(25) & vbTab & rsCov.Fields(41) & vbTab & rsCov.Fields(55)
								Case "LIF"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(9) & vbTab & rsCov.Fields(25) & vbTab & rsCov.Fields(41) & vbTab & rsCov.Fields(54)
								Case "LTD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(9) & vbTab & rsCov.Fields(25) & vbTab & rsCov.Fields(41) & vbTab & rsCov.Fields(53)
								Case "STD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(9) & vbTab & rsCov.Fields(25) & vbTab & rsCov.Fields(41) & vbTab & rsCov.Fields(56)
								Case "SUP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(9) & vbTab & rsCov.Fields(25) & vbTab & rsCov.Fields(41) & vbTab & rsCov.Fields(57)
								Case "SSP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(9) & vbTab & rsCov.Fields(25) & vbTab & rsCov.Fields(41) & vbTab & rsCov.Fields(58)
								Case "SDP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(9) & vbTab & rsCov.Fields(25) & vbTab & rsCov.Fields(41) & vbTab & rsCov.Fields(59)
								Case "SAD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(9) & vbTab & rsCov.Fields(25) & vbTab & rsCov.Fields(41) & vbTab & rsCov.Fields(60)
								Case "SM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(9) & vbTab & rsCov.Fields(25) & vbTab & rsCov.Fields(41) & vbTab & rsCov.Fields(61)
								Case "DM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(9) & vbTab & rsCov.Fields(25) & vbTab & rsCov.Fields(41) & vbTab & rsCov.Fields(62)
								Case Else
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(9) & vbTab & rsCov.Fields(25) & vbTab & rsCov.Fields(41)
								End Select
							End If
							If Len(Trim(rsCov.Fields(10))) > 0 Then
								Select Case rsCov.Fields(10)
								Case "ADD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(10) & vbTab & rsCov.Fields(26) & vbTab & rsCov.Fields(42) & vbTab & rsCov.Fields(55)
								Case "LIF"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(10) & vbTab & rsCov.Fields(26) & vbTab & rsCov.Fields(42) & vbTab & rsCov.Fields(54)
								Case "LTD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(10) & vbTab & rsCov.Fields(26) & vbTab & rsCov.Fields(42) & vbTab & rsCov.Fields(53)
								Case "STD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(10) & vbTab & rsCov.Fields(26) & vbTab & rsCov.Fields(42) & vbTab & rsCov.Fields(56)
								Case "SUP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(10) & vbTab & rsCov.Fields(26) & vbTab & rsCov.Fields(42) & vbTab & rsCov.Fields(57)
								Case "SSP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(10) & vbTab & rsCov.Fields(26) & vbTab & rsCov.Fields(42) & vbTab & rsCov.Fields(58)
								Case "SDP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(10) & vbTab & rsCov.Fields(26) & vbTab & rsCov.Fields(42) & vbTab & rsCov.Fields(59)
								Case "SAD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(10) & vbTab & rsCov.Fields(26) & vbTab & rsCov.Fields(42) & vbTab & rsCov.Fields(60)
								Case "SM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(10) & vbTab & rsCov.Fields(26) & vbTab & rsCov.Fields(42) & vbTab & rsCov.Fields(61)
								Case "DM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(10) & vbTab & rsCov.Fields(26) & vbTab & rsCov.Fields(42) & vbTab & rsCov.Fields(62)
								Case Else
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(10) & vbTab & rsCov.Fields(26) & vbTab & rsCov.Fields(42)
								End Select
							End If
							If Len(Trim(rsCov.Fields(11))) > 0 Then
								Select Case rsCov.Fields(11)
								Case "ADD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(11) & vbTab & rsCov.Fields(27) & vbTab & rsCov.Fields(43) & vbTab & rsCov.Fields(55)
								Case "LIF"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(11) & vbTab & rsCov.Fields(27) & vbTab & rsCov.Fields(43) & vbTab & rsCov.Fields(54)
								Case "LTD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(11) & vbTab & rsCov.Fields(27) & vbTab & rsCov.Fields(43) & vbTab & rsCov.Fields(53)
								Case "STD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(11) & vbTab & rsCov.Fields(27) & vbTab & rsCov.Fields(43) & vbTab & rsCov.Fields(56)
								Case "SUP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(11) & vbTab & rsCov.Fields(27) & vbTab & rsCov.Fields(43) & vbTab & rsCov.Fields(57)
								Case "SSP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(11) & vbTab & rsCov.Fields(27) & vbTab & rsCov.Fields(43) & vbTab & rsCov.Fields(58)
								Case "SDP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(11) & vbTab & rsCov.Fields(27) & vbTab & rsCov.Fields(43) & vbTab & rsCov.Fields(59)
								Case "SAD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(11) & vbTab & rsCov.Fields(27) & vbTab & rsCov.Fields(43) & vbTab & rsCov.Fields(60)
								Case "SM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(11) & vbTab & rsCov.Fields(27) & vbTab & rsCov.Fields(43) & vbTab & rsCov.Fields(61)
								Case "DM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(11) & vbTab & rsCov.Fields(27) & vbTab & rsCov.Fields(43) & vbTab & rsCov.Fields(62)
								Case Else
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(11) & vbTab & rsCov.Fields(27) & vbTab & rsCov.Fields(43)
								End Select
							End If
							If Len(Trim(rsCov.Fields(12))) > 0 Then
								Select Case rsCov.Fields(12)
								Case "ADD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(12) & vbTab & rsCov.Fields(28) & vbTab & rsCov.Fields(44) & vbTab & rsCov.Fields(55)
								Case "LIF"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(12) & vbTab & rsCov.Fields(28) & vbTab & rsCov.Fields(44) & vbTab & rsCov.Fields(54)
								Case "LTD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(12) & vbTab & rsCov.Fields(28) & vbTab & rsCov.Fields(44) & vbTab & rsCov.Fields(53)
								Case "STD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(12) & vbTab & rsCov.Fields(28) & vbTab & rsCov.Fields(44) & vbTab & rsCov.Fields(56)
								Case "SUP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(12) & vbTab & rsCov.Fields(28) & vbTab & rsCov.Fields(44) & vbTab & rsCov.Fields(57)
								Case "SSP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(12) & vbTab & rsCov.Fields(28) & vbTab & rsCov.Fields(44) & vbTab & rsCov.Fields(58)
								Case "SDP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(12) & vbTab & rsCov.Fields(28) & vbTab & rsCov.Fields(44) & vbTab & rsCov.Fields(59)
								Case "SAD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(12) & vbTab & rsCov.Fields(28) & vbTab & rsCov.Fields(44) & vbTab & rsCov.Fields(60)
								Case "SM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(12) & vbTab & rsCov.Fields(28) & vbTab & rsCov.Fields(44) & vbTab & rsCov.Fields(61)
								Case "DM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(12) & vbTab & rsCov.Fields(28) & vbTab & rsCov.Fields(44) & vbTab & rsCov.Fields(62)
								Case Else
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(12) & vbTab & rsCov.Fields(28) & vbTab & rsCov.Fields(44)
								End Select
							End If
							If Len(Trim(rsCov.Fields(13))) > 0 Then
								Select Case rsCov.Fields(13)
								Case "ADD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(13) & vbTab & rsCov.Fields(29) & vbTab & rsCov.Fields(45) & vbTab & rsCov.Fields(55)
								Case "LIF"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(13) & vbTab & rsCov.Fields(29) & vbTab & rsCov.Fields(45) & vbTab & rsCov.Fields(54)
								Case "LTD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(13) & vbTab & rsCov.Fields(29) & vbTab & rsCov.Fields(45) & vbTab & rsCov.Fields(53)
								Case "STD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(13) & vbTab & rsCov.Fields(29) & vbTab & rsCov.Fields(45) & vbTab & rsCov.Fields(56)
								Case "SUP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(13) & vbTab & rsCov.Fields(29) & vbTab & rsCov.Fields(45) & vbTab & rsCov.Fields(57)
								Case "SSP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(13) & vbTab & rsCov.Fields(29) & vbTab & rsCov.Fields(45) & vbTab & rsCov.Fields(58)
								Case "SDP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(13) & vbTab & rsCov.Fields(29) & vbTab & rsCov.Fields(45) & vbTab & rsCov.Fields(59)
								Case "SAD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(13) & vbTab & rsCov.Fields(29) & vbTab & rsCov.Fields(45) & vbTab & rsCov.Fields(60)
								Case "SM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(13) & vbTab & rsCov.Fields(29) & vbTab & rsCov.Fields(45) & vbTab & rsCov.Fields(61)
								Case "DM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(13) & vbTab & rsCov.Fields(29) & vbTab & rsCov.Fields(45) & vbTab & rsCov.Fields(62)
								Case Else
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(13) & vbTab & rsCov.Fields(29) & vbTab & rsCov.Fields(45)
								End Select
							End If
							If Len(Trim(rsCov.Fields(14))) > 0 Then
								Select Case rsCov.Fields(14)
								Case "ADD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(14) & vbTab & rsCov.Fields(30) & vbTab & rsCov.Fields(46) & vbTab & rsCov.Fields(55)
								Case "LIF"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(14) & vbTab & rsCov.Fields(30) & vbTab & rsCov.Fields(46) & vbTab & rsCov.Fields(54)
								Case "LTD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(14) & vbTab & rsCov.Fields(30) & vbTab & rsCov.Fields(46) & vbTab & rsCov.Fields(53)
								Case "STD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(14) & vbTab & rsCov.Fields(30) & vbTab & rsCov.Fields(46) & vbTab & rsCov.Fields(56)
								Case "SUP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(14) & vbTab & rsCov.Fields(30) & vbTab & rsCov.Fields(46) & vbTab & rsCov.Fields(57)
								Case "SSP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(14) & vbTab & rsCov.Fields(30) & vbTab & rsCov.Fields(46) & vbTab & rsCov.Fields(58)
								Case "SDP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(14) & vbTab & rsCov.Fields(30) & vbTab & rsCov.Fields(46) & vbTab & rsCov.Fields(59)
								Case "SAD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(14) & vbTab & rsCov.Fields(30) & vbTab & rsCov.Fields(46) & vbTab & rsCov.Fields(60)
								Case "SM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(14) & vbTab & rsCov.Fields(30) & vbTab & rsCov.Fields(46) & vbTab & rsCov.Fields(61)
								Case "DM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(14) & vbTab & rsCov.Fields(30) & vbTab & rsCov.Fields(46) & vbTab & rsCov.Fields(62)
								Case Else
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(14) & vbTab & rsCov.Fields(30) & vbTab & rsCov.Fields(46)
								End Select
							End If
							If Len(Trim(rsCov.Fields(15))) > 0 Then
								Select Case rsCov.Fields(15)
								Case "ADD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(15) & vbTab & rsCov.Fields(31) & vbTab & rsCov.Fields(47) & vbTab & rsCov.Fields(55)
								Case "LIF"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(15) & vbTab & rsCov.Fields(31) & vbTab & rsCov.Fields(47) & vbTab & rsCov.Fields(54)
								Case "LTD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(15) & vbTab & rsCov.Fields(31) & vbTab & rsCov.Fields(47) & vbTab & rsCov.Fields(53)
								Case "STD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(15) & vbTab & rsCov.Fields(31) & vbTab & rsCov.Fields(47) & vbTab & rsCov.Fields(56)
								Case "SUP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(15) & vbTab & rsCov.Fields(31) & vbTab & rsCov.Fields(47) & vbTab & rsCov.Fields(57)
								Case "SSP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(15) & vbTab & rsCov.Fields(31) & vbTab & rsCov.Fields(47) & vbTab & rsCov.Fields(58)
								Case "SDP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(15) & vbTab & rsCov.Fields(31) & vbTab & rsCov.Fields(47) & vbTab & rsCov.Fields(59)
								Case "SAD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(15) & vbTab & rsCov.Fields(31) & vbTab & rsCov.Fields(47) & vbTab & rsCov.Fields(60)
								Case "SM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(15) & vbTab & rsCov.Fields(31) & vbTab & rsCov.Fields(47) & vbTab & rsCov.Fields(61)
								Case "DM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(15) & vbTab & rsCov.Fields(31) & vbTab & rsCov.Fields(47) & vbTab & rsCov.Fields(62)
								Case Else
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(15) & vbTab & rsCov.Fields(31) & vbTab & rsCov.Fields(47)
								End Select
							End If
							If Len(Trim(rsCov.Fields(16))) > 0 Then
								Select Case rsCov.Fields(16)
								Case "ADD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(16) & vbTab & rsCov.Fields(32) & vbTab & rsCov.Fields(48) & vbTab & rsCov.Fields(55)
								Case "LIF"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(16) & vbTab & rsCov.Fields(32) & vbTab & rsCov.Fields(48) & vbTab & rsCov.Fields(54)
								Case "LTD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(16) & vbTab & rsCov.Fields(32) & vbTab & rsCov.Fields(48) & vbTab & rsCov.Fields(53)
								Case "STD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(16) & vbTab & rsCov.Fields(32) & vbTab & rsCov.Fields(48) & vbTab & rsCov.Fields(56)
								Case "SUP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(16) & vbTab & rsCov.Fields(32) & vbTab & rsCov.Fields(48) & vbTab & rsCov.Fields(57)
								Case "SSP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(16) & vbTab & rsCov.Fields(32) & vbTab & rsCov.Fields(48) & vbTab & rsCov.Fields(58)
								Case "SDP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(16) & vbTab & rsCov.Fields(32) & vbTab & rsCov.Fields(48) & vbTab & rsCov.Fields(59)
								Case "SAD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(16) & vbTab & rsCov.Fields(32) & vbTab & rsCov.Fields(48) & vbTab & rsCov.Fields(60)
								Case "SM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(16) & vbTab & rsCov.Fields(32) & vbTab & rsCov.Fields(48) & vbTab & rsCov.Fields(61)
								Case "DM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(16) & vbTab & rsCov.Fields(32) & vbTab & rsCov.Fields(48) & vbTab & rsCov.Fields(62)
								Case Else
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(16) & vbTab & rsCov.Fields(32) & vbTab & rsCov.Fields(48)
								End Select
							End If
							If Len(Trim(rsCov.Fields(17))) > 0 Then
								Select Case rsCov.Fields(17)
								Case "ADD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(17) & vbTab & rsCov.Fields(33) & vbTab & rsCov.Fields(49) & vbTab & rsCov.Fields(55)
								Case "LIF"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(17) & vbTab & rsCov.Fields(33) & vbTab & rsCov.Fields(49) & vbTab & rsCov.Fields(54)
								Case "LTD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(17) & vbTab & rsCov.Fields(33) & vbTab & rsCov.Fields(49) & vbTab & rsCov.Fields(53)
								Case "STD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(17) & vbTab & rsCov.Fields(33) & vbTab & rsCov.Fields(49) & vbTab & rsCov.Fields(56)
								Case "SUP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(17) & vbTab & rsCov.Fields(33) & vbTab & rsCov.Fields(49) & vbTab & rsCov.Fields(57)
								Case "SSP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(17) & vbTab & rsCov.Fields(33) & vbTab & rsCov.Fields(49) & vbTab & rsCov.Fields(58)
								Case "SDP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(17) & vbTab & rsCov.Fields(33) & vbTab & rsCov.Fields(49) & vbTab & rsCov.Fields(59)
								Case "SAD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(17) & vbTab & rsCov.Fields(33) & vbTab & rsCov.Fields(49) & vbTab & rsCov.Fields(60)
								Case "SM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(17) & vbTab & rsCov.Fields(33) & vbTab & rsCov.Fields(49) & vbTab & rsCov.Fields(61)
								Case "DM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(17) & vbTab & rsCov.Fields(33) & vbTab & rsCov.Fields(49) & vbTab & rsCov.Fields(62)
								Case Else
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(17) & vbTab & rsCov.Fields(33) & vbTab & rsCov.Fields(49)
								End Select
							End If
							If Len(Trim(rsCov.Fields(18))) > 0 Then
								Select Case rsCov.Fields(18)
								Case "ADD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(18) & vbTab & rsCov.Fields(34) & vbTab & rsCov.Fields(50) & vbTab & rsCov.Fields(55)
								Case "LIF"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(18) & vbTab & rsCov.Fields(34) & vbTab & rsCov.Fields(50) & vbTab & rsCov.Fields(54)
								Case "LTD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(18) & vbTab & rsCov.Fields(34) & vbTab & rsCov.Fields(50) & vbTab & rsCov.Fields(53)
								Case "STD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(18) & vbTab & rsCov.Fields(34) & vbTab & rsCov.Fields(50) & vbTab & rsCov.Fields(56)
								Case "SUP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(18) & vbTab & rsCov.Fields(34) & vbTab & rsCov.Fields(50) & vbTab & rsCov.Fields(57)
								Case "SSP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(18) & vbTab & rsCov.Fields(34) & vbTab & rsCov.Fields(50) & vbTab & rsCov.Fields(58)
								Case "SDP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(18) & vbTab & rsCov.Fields(34) & vbTab & rsCov.Fields(50) & vbTab & rsCov.Fields(59)
								Case "SAD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(18) & vbTab & rsCov.Fields(34) & vbTab & rsCov.Fields(50) & vbTab & rsCov.Fields(60)
								Case "SM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(18) & vbTab & rsCov.Fields(34) & vbTab & rsCov.Fields(50) & vbTab & rsCov.Fields(61)
								Case "DM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(18) & vbTab & rsCov.Fields(34) & vbTab & rsCov.Fields(50) & vbTab & rsCov.Fields(62)
								Case Else
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(18) & vbTab & rsCov.Fields(34) & vbTab & rsCov.Fields(50)
								End Select
							End If
							If Len(Trim(rsCov.Fields(19))) > 0 Then
								Select Case rsCov.Fields(19)
								Case "ADD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(19) & vbTab & rsCov.Fields(35) & vbTab & rsCov.Fields(51) & vbTab & rsCov.Fields(55)
								Case "LIF"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(19) & vbTab & rsCov.Fields(35) & vbTab & rsCov.Fields(51) & vbTab & rsCov.Fields(54)
								Case "LTD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(19) & vbTab & rsCov.Fields(35) & vbTab & rsCov.Fields(51) & vbTab & rsCov.Fields(53)
								Case "STD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(19) & vbTab & rsCov.Fields(35) & vbTab & rsCov.Fields(51) & vbTab & rsCov.Fields(56)
								Case "SUP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(19) & vbTab & rsCov.Fields(35) & vbTab & rsCov.Fields(51) & vbTab & rsCov.Fields(57)
								Case "SSP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(19) & vbTab & rsCov.Fields(35) & vbTab & rsCov.Fields(51) & vbTab & rsCov.Fields(58)
								Case "SDP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(19) & vbTab & rsCov.Fields(35) & vbTab & rsCov.Fields(51) & vbTab & rsCov.Fields(59)
								Case "SAD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(19) & vbTab & rsCov.Fields(35) & vbTab & rsCov.Fields(51) & vbTab & rsCov.Fields(60)
								Case "SM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(19) & vbTab & rsCov.Fields(35) & vbTab & rsCov.Fields(51) & vbTab & rsCov.Fields(61)
								Case "DM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(19) & vbTab & rsCov.Fields(35) & vbTab & rsCov.Fields(51) & vbTab & rsCov.Fields(62)
								Case Else
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(19) & vbTab & rsCov.Fields(35) & vbTab & rsCov.Fields(51)
								End Select
							End If						
							If Len(Trim(rsCov.Fields(20))) > 0 Then
								Select Case rsCov.Fields(20)
								Case "ADD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(20) & vbTab & rsCov.Fields(36) & vbTab & rsCov.Fields(52) & vbTab & rsCov.Fields(55)
								Case "LIF"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(20) & vbTab & rsCov.Fields(36) & vbTab & rsCov.Fields(52) & vbTab & rsCov.Fields(54)
								Case "LTD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(20) & vbTab & rsCov.Fields(36) & vbTab & rsCov.Fields(52) & vbTab & rsCov.Fields(53)
								Case "STD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(20) & vbTab & rsCov.Fields(36) & vbTab & rsCov.Fields(52) & vbTab & rsCov.Fields(56)
								Case "SUP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(20) & vbTab & rsCov.Fields(36) & vbTab & rsCov.Fields(52) & vbTab & rsCov.Fields(57)
								Case "SSP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(20) & vbTab & rsCov.Fields(36) & vbTab & rsCov.Fields(52) & vbTab & rsCov.Fields(58)
								Case "SDP"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(20) & vbTab & rsCov.Fields(36) & vbTab & rsCov.Fields(52) & vbTab & rsCov.Fields(59)
								Case "SAD"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(20) & vbTab & rsCov.Fields(36) & vbTab & rsCov.Fields(52) & vbTab & rsCov.Fields(60)
								Case "SM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(20) & vbTab & rsCov.Fields(36) & vbTab & rsCov.Fields(52) & vbTab & rsCov.Fields(61)
								Case "DM2"
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(20) & vbTab & rsCov.Fields(36) & vbTab & rsCov.Fields(52) & vbTab & rsCov.Fields(62)
								Case Else
								objCovTextFile.Writeline sID & vbTab & rsCov.Fields(20) & vbTab & rsCov.Fields(36) & vbTab & rsCov.Fields(52)
								End Select
							End If	
						lstID = rsCov.Fields(4)
						rsCov.MoveNext
						End If

			Loop
			
			'==================================
			'Writes Eligibility record to file.
			'==================================
			sEligRec = ""
			For i = 0 to rsElig.Fields.Count - 1
						sEligRec = sEligRec & rsElig.Fields(i) & chr(9)
			Next
			'sEligRec = sEligRec & "26" & chr(9) & rsElig.Fields(23) & chr(9) & rsElig.Fields(24)
			objEligTextFile.Writeline sEligRec
			lstID = rsElig.Fields(0)                         
			rsElig.MoveNext
	End If
	
Loop

objEligTextFile.Close
objCovTextFile.Close

EndTime = Timer()

sTimeTaken = FormatNumber(EndTime - StartTime, 2)

MsgBox "Done. " & counter & " records processed in " & sTimeTaken & " seconds.", vbInformation

Set rsElig = Nothing
Set rsCov = Nothing
Set objEligTextFile = Nothing
Set objCovTextFile = Nothing
Set objFSO = Nothing
Set objConn = Nothing 
'=*=*=*=*=*=*=*=*=*=*=*=
' Coded By Assaf Miron 
' http://assaf.miron.googlepages.com
' Date : 03/02/08
'=*=*=*=*=*=*=*=*=*=*=*=
'*================================================*
'* This Script is Mainly used to Export Users and Computers from the Production Domain
'* and Import Them into a Test Domain.
'* This Script Creates all the OU Structure and Resets all the Users Password to a Default Password
'* The Script Imports the Exported Users and Computers to Their Disabled State 
'* (If the User is Disabled it is Imported Disabled)
'* The Script Saves all of its actions to a Log File
'*================================================*
'================================================
' Consts
'================================================
' File Consts
Const ForReading = 1 
Const ForWriting = 2
Const ForAppending = 8
' Excel Consts
Const xlCenter = -4108
Const xlSolid = 1
' Active Directory Consts
Const ADS_SCOPE_SUBTREE = 2
Const ADS_UF_ACCOUNTDISABLE = &h00002
Const DefPassword = "Password1"

' Log File Name and Path
Const LogFile = "c:\CreateOU-LogFile.txt"

'================================================
' Sets
'================================================
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objExcel = CreateObject("Excel.Application")
Set objOUExcel = CreateObject("Excel.Application")
Set objRootDSE = GetObject ("LDAP://rootDSE")

'================================================
' Dims and Variables
'================================================
Dim strDomainLdap, objRootDSE
Dim arrOU
Dim FileLoc

strDomainLdap  = objRootDSE.Get("defaultNamingContext")
ouLDAP = "ou=Test," & strDomainLdap
strPrincipleName = "@" & Replace(strDomainLdap,"DC=","")
strPrincipleName = Replace(strPrincipleName,",",".")

'================================================
' Functions
'================================================
Function ErrorMSG()
' This Function checks the Error Number and if it is not 0 then it outputs the Error in a MsgBox
	If Not Err.Number = 0 Then
		MsgBox "Error # " & CStr(Err.Number) & " " & Err.Description
		objFile.WriteLine "Error # " & CStr(Err.Number) & " " & Err.Description
		Err.Clear    ' Clear the error.
	End If
End Function

Function CheckComma(strAccount,OUName)
' This Function Checks the LDAP Syntax for Extra Commas
	' Check the OUName Last Char
	chkComma = Mid(OUName,Len(OUName),1)
	' Acording to the strAccount and chkComma Values Format and Return the LDAP String
	If (strAccount = "") Then
		If chkComma = "," Then
			CheckComma = "LDAP://" & OUName & ouLDAP
		Else
			CheckComma = "LDAP://" & OUName & "," & ouLDAP
		End If
	Else
		If chkComma = "," Then
			CheckComma = "LDAP://cn=" & strAccount & "," & OUName & ouLDAP
		Else
			CheckComma = "LDAP://cn=" & strAccount & "," & OUName & "," & ouLDAP
		End If
	End If
End Function

Function CheckOUName(OUName)
' This Function Checks the OU Name for Special Charecters 
	NewOUName = OUName
	If instr(OUName,"\" & chr(34)) Then
		CheckOUName = NewOUName
		Exit Function
	End If
	If instr(OUName,chr(34)) Then
		NewOUName = Replace(OUName,chr(34),"\" & chr(34))
		CheckOUName = NewOUName
		Exit Function
	End If
	If instr(OUName,"/") Then
		NewOUName = Replace(OUName,"/","\/")
		CheckOUName = NewOUName
		Exit Function
	End If
	CheckOUName = NewOUName
End Function

Function OpenFile()
' This Function Opens a Common Dialog for the User to Select a File to Open
' Filter the Files to Excel and CSV Files Only
	Set objDialog = CreateObject("UserAccounts.CommonDialog")
	
	objDialog.Filter = "Excel Files|*.xls|CSV Files|*.csv"
	objDialog.FilterIndex = 1
	objDialog.InitialDir = "C:\"
	intResult = objDialog.ShowOpen
	 
	If intResult = 0 Then
	    Wscript.Quit
	Else
	    FileLoc = objDialog.FileName
	End If
	objFile.WriteLine "Opening File : " & FileLoc
	objFile.WriteLine
	
	OpenFile = FileLoc
End Function

'================================================
' Subs
'================================================
Sub ExcelHeaders()
'Create Excel Headers and color them in gray

Set objRange = objExcel.Range("A1","H1")
objRange.Font.Size = 12
objRange.Interior.ColorIndex=15

objexcel.cells(1,1)="OU Name"
objexcel.cells(1,2)="OU Description"
objexcel.cells(1,3)="Object Type (User/Computer)"
objExcel.cells(1,4)="Object Account Name"
objExcel.cells(1,5)="First Name"
objExcel.cells(1,6)="Surname"
objExcel.cells(1,7)="Display Name"
objExcel.Cells(1,8)="Disabled State"

End Sub

Sub CenterCells()
'Format the Cells and Headers
Letters = "A,B,C,D,E,F,G,H"
arrLetters = Split(Letters,",")
For I= 0 to Ubound(arrLetters)
	Ltr = arrLetters(i)
	set objRange = objExcel.Range(Ltr & "1")
	objRange.Activate
	With objExcel.Selection
	    .HorizontalAlignment = xlCenter
	    .WrapText = False
	    .Orientation = 0
	    .AddIndent = False
	    .IndentLevel = 0
	    .ShrinkToFit = False
	    .MergeCells = False
	    .AutoFilter
	End With

	With objExcel.Selection.Interior
	    .ColorIndex = 15
	    .Pattern = xlSolid
	End With
	' Split the First Row so Scrolling would be Easy
	objExcel.ActiveWindow.SplitRow = 1
	objExcel.ActiveWindow.FreezePanes = True
	'Auto Fit the Cells
	Set ObjRange = objExcel.ActiveCell.EntireColumn
	objRange.AutoFit()
Next
End Sub

Sub OUInfo(SecOU,MainOU,OUDesc)
' This Sub Collects and Exports OU Information
' It Exports all the Users and Computers Information into an Excel File
If Err.Number <> 0 Then
	ErrorMSG
End If
' Check if the OU's Name is Valid
MainOU = CheckOUName(MainOU)
SecOU = CheckOUName(SecOU)
' Connect to the OU
Set objContainer = GetObject("GC://" & SecOU & "," & MainOU & "," & ouLdap)
strOUName = MainOU
strOUDesc = SecOU
objFile.WriteLine "Exporting Users and Computer from OU " & strOUName
For Each objMember in objContainer
	objExcel.Cells(intRow,1) = strOUName
	objExcel.Cells(intRow,2) = strOUDesc
	If objMember.Class = "user" Then
	' Export User Information
		objExcel.cells(introw,3) = "User"
		objExcel.cells(introw,4) = objMember.sAMAccountName
		objExcel.cells(introw,5) = objMember.givenname
		objExcel.cells(introw,6) = objMember.sn
		objExcel.cells(introw,7) = objMember.displayName
		intUAC = objMember.Get("userAccountControl")
		introw=introw+1
  	ElseIf objMember.Class = "organizationalUnit" Then
  	' If Objcet is an OU the Call this sub Again with the new Secondry OU
		tmpOU = SecOU & "," & MainOU
		OUInfo objMember.Name,tmpOU,objMember.Description
	ElseIf objMember.Class = "computer" Then
	' Export Computer Information
		objExcel.cells(introw,3) = "Computer"
		objExcel.cells(introw,4) = objMember.sAMAccountName
		introw=introw+1
	End If
	' Check the Disabled State of The Object
		If intUAC And ADS_UF_ACCOUNTDISABLE Then
			objexcel.cells(introw,8)= "True"
		End If
Next
End Sub

Sub CreateOU(MainOU,SecOU)
' This Sub Creates the OU Hierarchy
	Msg = "Creating OU " + SecOU
	objFile.WriteLine "Creating Sub OU "
	Err = objFile.WriteLine("-----" & Msg)
	If (MainOU <> " ") Then
		If InStr(MainOU,",") = 0 Then
			MainOU = MainOU & ","
		End If
		Set objDomain = GetObject(CheckComma("",MainOU))
	Else
		Set objDomain = GetObject("LDAP://" & ouLDAP)
	End If
	Set objOU = objDomain.Create("organizationalUnit", SecOU)
  	objOU.SetInfo
End Sub

Sub CheckOUExists(MainOU,SecOU)
' This Sub Checks if the MainOU or Secondary OU Exists
' If the OU's doesnt Exists it Creates the OU Hierarchy
	OUPath = ""
	' If the MainOU contains a Path of more than one OU
	If InStr(MainOU,",") Then
		arrOU = Split(MainOU,",")
		For I = UBound(arrOU) to 1 step -1
			OUPath = arrOU(i) & "," & OUPath
			objFile.WriteLine "Connecting to : GC://" & OUPath & ouLDAP
			Set objContainer = GetObject("GC://" & OUPath & ouLDAP)
			objContainer.Filter = Array("organizationalUnit")
			If IsEmpty(objContainer) = True Then
				objFile.WriteLine "No OU's in Container " & objContainer
				CreateOU OUPath,arrOU(i-1)
			Else
				ToCreate = 0
				For Each OU in objContainer
					If UCase(OU.Name) = UCase(arrOU(i-1)) Then
						objFile.WriteLine "OU Exists " & OU.Name
						ToCreate = 1
						Exit For					
					End If
				Next
				If ToCreate = 0 Then
					CreateOU OUPath,arrOU(i-1)
				End If
			End If
		Next
		OUPath = arrOU(0) & "," & OUPath
		ToCreate = 0
		Set objContainer = GetObject("GC://" & OUPath & ouLDAP)
		objContainer.Filter = Array("organizationalUnit")
		For Each OU in objContainer
			If UCase(OU.Name) = UCase(SecOU) Then
				objFile.WriteLine "OU Exists " & OU.Name
				ToCreate = 1
				Exit For					
			End If
		Next
		If ToCreate = 0 Then
			CreateOU OUPath,SecOU
		End If
	
		Set objContainer = GetObject("GC://" & MainOU & "," & ouLDAP)
		objContainer.Filter = Array("organizationalUnit")
		If IsEmpty(objContainer) = True Then
			wscript.echo "IsEmpty"
			objFile.WriteLine "No OU's in Container " & objContainer
			CreateOU MainOU,SecOU
		Else
			For Each OU in objContainer
				If OU.Name = MainOU Then
					objFile.WriteLine "OU Exists " & OU.Name
					Exit For					
				End If
			Next
		End If
	Else
	' If the MainOU does not Contains a Path of OU's
		Set objContainer = GetObject("GC://" & ouLDAP)
		objContainer.Filter = Array("organizationalUnit")
		ToCreate = 0
		For Each OU in ObjContainer
			If UCase(OU.Name) = UCase(MainOU) Then
				objFile.WriteLine "OU Exists " & OU.Name
				ToCreate = 1
				Exit For
			Else
				ToCreate = 0
			End If
		Next
		If ToCreate = 0 Then
			CreateOU " ",MainOU
		End If
	End If
	Set objContainer = GetObject("GC://" & MainOU & "," & ouLDAP)
	objContainer.Filter = Array("organizationalUnit")
	ToCreate = 0
	For Each OU in objContainer
		If UCase(OU.Name) = UCase(SecOU) Then
			objFile.WriteLine "OU Exists " & OU.Name
			ToCreate = 1
			Exit For
		Else
			ToCreate = 0
		End If
	Next
	If ToCreate = 0 Then
		CreateOU MainOU,SecOU
	End If
End Sub

Sub CreateADObject(AccType,AccountName,Name,SName,Description,UPN,OU,Disabled)
' This Sub Creates an AD Object acourding to it's Type, Name and Disabled State
On Error Resume Next
	Set objOU = GetObject(CheckComma("",OU))
	objFile.WriteLine "Creating " & AccType & " (" & AccountName & ") in OU : " & OU
	objOU.Filter = Array(AccType)
	ToCreate = 0
	For Each oUser in objOU
	temp = oUser.sAMAccountName
		If UCase(oUser.sAMAccountName) = UCase(AccountName) Then
			objFile.WriteLine AccType & " " & AccountName & " Exists"
			ToCreate = 1
			Exit For
		Else
			ToCreate = 0
		End If
	Next
	' Error Handling, you cant Pass a Null Value when Creating a User Account
	' If the Name, Surname or Description is Null then Convert them to a Space Value
	If SName = "" Then
		SName = " "
	End If
	If Name = "" Then
		Name = " "
	End If
	If Description = "" Then
		Description = " "
	End If
	If ToCreate = 0 Then
		' Create the object, if it is a Computer remove the $ sign
		Set objAccount = objOU.Create(AccType, "cn=" & Replace(AccountName,"$",""))
		objAccount.Put "sAMAccountName", CStr(AccountName)
		If LCase(AccType) = LCase("user") Then
			objAccount.Put "givenName",CStr(Name)
			objAccount.Put "sn",CStr(SName)
			objAccount.Put "displayName",CStr(Description)
			objAccount.Put "userPrincipalName", UPN
		End If
		err = objAccount.SetInfo
		ErrorMSG
		objFile.WriteLine "Creating " & AccountName
	End If
	' Reset the User's Password
	Set objAccount = GetObject(CheckComma(Replace(AccountName,"$",""),OU))
	If AccType = "user" Then
		objAccount.SetPassword DefPassword
	End If
	' Configure the User or Computer Acourding to it's Disabled State
	objAccount.AccountDisabled = Disabled
 	Err = objAccount.SetInfo
 	ErrorMSG
End Sub

Sub ChangeUPN(strAccount,OU)
On Error Resume Next
' This Sub Changes the Users UPN and mail Attribute
' This is Needed doe to an Unexplainable Bug that cant assign the UPN when Creating the User Account
	Set User = GetObject(CheckComma(strAccount,OU))
	strsAMAccountName=User.sAMAccountName
	User.Put "userPrincipalName", strsAMAccountName & strPrincipleName 
 	User.Put "mail", strsAMAccountName & strPrincipleName 
	User.SetInfo
End Sub

Sub ExportADObjects(MainOU,intWorksheet)
' This Sub Exports all the Users and Computer accounts in the MainOU OU
' The intWorksheet indicates on what Excel WorkSheet to Work on and Export the information
	If instr(MainOU,"-") Then
		arrOUs = Split(MainOu,"-")
		If Ubound(arrOUs) = 1 Then
			MainOU = "ou=" & arrOUs(1) & "," & "ou=" & arrOUs(0)
		End If
	Else
		MainOU = "ou=" & MainOU
	End If
' Check if the OU Name is Valid
	MainOU = CheckOUName(MainOU)
' Connect to the MainOU OU	
	Set objContainer = GetObject _
	  ("GC://" & MainOU & "," & ouLdap)
' Open the Excel WorkSheet and Make It Active	
	Set ObjWorkSheet=ObjExcel.WorkSheets(intWorksheet)
	ObjWorkSheet.Activate
	If objContainer.Description = "" Then
		objWorkSheet.Name = Replace(MainOU,"\","")
	Else
		objWorkSheet.Name = objContainer.Description
	End If
	'Create Nice Headers
	ExcelHeaders
	introw = 2
	If Len(MainOU) > 12 Then
		if instr(MainOU,",") Then
			arrOU = Split(MainOU,",")
			MainOU = arrOU(1)
			SecOU = arrOU(0)
			OUInfo SecOU,MainOU,SecOU
		Else
			For Each OU in objContainer
			' Export Information on all the Sub OU in the MainOU OU
				If (OU.Class <> "group") AND (OU.Class <> "user") Then
					objexcel.cells(introw,1) = OU.Name
					SecOU = OU.Name
					objexcel.cells(introw,2) = OU.Description
					OUInfo SecOU,MainOU,OU.Description
				End If
			Next
		End If
	Else
		For Each OU in objContainer
			If (OU.Class <> "group") AND (OU.Class <> "user") And (ou.Class = "organizationalUnit") Then
				SecOU = OU.Name
				OUInfo SecOU,MainOU,OU.Description
			End If
		Next
	End If
End Sub

Sub ImportADObjects()
' This Sub Imports the AD Object Listed in the Excel File Opened
intWB = 0
' Count the WorkSheets in the Opend File
For Each WS in objOUExcel.WorkSheets
	intWS = intWS + 1
Next

objFile.WriteLine "There are " & intWS & " WorkSheets in this File"

For w = 1 to intWS
	introw=2
	set ObjWorkSheet=ObjOUExcel.WorkSheets(w)
	' Activate Each WorkSheet and Read the Data from it
	ObjWorkSheet.Activate
	objFile.WriteLine "Reading from WorkSheet " & ObjWorkSheet.Name
	objfile.WriteLine
	Do Until objOUExcel.Cells(intRow,1).Value = ""
			MainOU = objOUExcel.Cells(intRow,1).Value
			SecOU = objOUExcel.Cells(intRow,2).Value
			objFile.WriteLine "Reading : " & MainOU & "," & SecOU & "."
			If InStr(MainOU,",") Then
				arrOU = Split(MainOU,",")
				Set objContainer = GetObject("GC://" & ouLDAP)
				objContainer.Filter = Array("organizationalUnit")
				ToCreate = 0
				' Check if the OU Needs to be Created
				For Each OU in ObjContainer
					If UCase(OU.Name) = UCase(arrOU(UBound(arrOU))) Then
						objFile.WriteLine "OU Exists " & OU.Name
						ToCreate = 1
						Exit For
					Else
						ToCreate = 0
						' If the ToCreate equals 0 Then Create the OU
					End If
				Next
				If ToCreate = 0 Then
					CreateOU " ",arrOU(UBound(arrOU))
				End If
			End If
			' Validate that the OU Exists
			CheckOUExists MainOU,SecOU
			' Collect the Object Data
			AccountType = objOUExcel.Cells(intRow,3).Value
			AccountID = objOUExcel.Cells(intRow,4).Value
			AccountFName = objOUExcel.Cells(intRow,5).Value
			AccountSName = objOUExcel.Cells(intRow,6).Value
			AccountDesc = objOUExcel.Cells(intRow,7).Value
			AccountDisabled = objOUExcel.Cells(intRow,8).Value
			UPN = UserID & strPrincipleName
			' Create the AD Object
			CreateADObject AccountType,AccountID,AccountFName,AccountSName,AccountDesc,UPN,SecOU & "," & MainOU,AccountDisabled
			If LCase(AccountType) = LCase("User") Then
			' If the Account is a User then Change it's UPN
				ChangeUPN AccountID, SecOU & "," & MainOU
			End If
			intRow = intRow + 1
	Loop
Next

End Sub

Sub Cleanup()
' This Sub Cleans up the Script
' Ends the Log File and Quits all Remaining Excel Files
objFile.WriteLine
objFile.WriteLine "Script Ended (" & Now & ")"
objExcel.Quit
objOUExcel.Quit
objFile.close
End Sub

Sub ShowHelp(ExtraMsg)
' This Sub Shows a Message box that show the Syntax and Usage of the Script
strUsageMsg = "Try Running this Script with an Argument of Import or Export" & vbCrLf
strUsageMsg = strUsageMsg & vbCrLf & "Example : Wscript " & WScript.ScriptFullName & " /Import" 
strUsageMsg = strUsageMsg & vbCrLf & "Example : Wscript " & WScript.ScriptFullName & " /Export"
If ExtraMsg <> "" Then
	If ExtraMsg = "Help" Then
		strMSG = "About:" & vbNewLine & "Created by : Assaf Miron" & vbNewLine & "Date : 03/02/2008"
		strMSG = strUsageMsg & vbNewLine & vbNewLine & strMSG
		MsgBox strMSG,vbInformation + vbOKOnly,"Import/Export AD Objects"
	Else
		strMSG = ExtraMsg & vbCrLf & vbCrLf & strUsageMsg
		MsgBox strMSG,vbExclamation + vbOKOnly,"Import/Export AD Objects"
	End If
Else
	strMSG = strUsageMsg
	MsgBox strMSG,vbInformation + vbOKOnly,"Import/Export AD Objects"
End If
' Close Open Instances and Free Files
Cleanup
WScript.Quit
End Sub

' Main Code
' Check the Wscript Arguments 
' If no Arguments assinged then Show the Help Usage
If WScript.Arguments.Count = 0 Then
	ShowHelp ""
Else
' If more than one Arguments assinged then Show the Help Usage
	If WScript.Arguments.Count > 1 Then
		ShowHelp "To Many Arguments"
	End If
' If one Argument assinged then assign it to the Action Variable
	If WScript.Arguments.Count = 1 Then
		Action = WScript.Arguments(0)
	End If
End If

' Create the Log File
Set objFile = objFSO.CreateTextFile(LogFile,ForWriting)
objFile.WriteLine
objFile.WriteLine "Log Started Now : " & NOW

introw=2
' Check First if the Assigned Action Variable is Correct and Then Open the Excel File to Read
If (LCase(Action)=LCase("/Import")) Or (LCase(Action)=LCase("/Export")) Then
	Set objWorkbook = objOUExcel.Workbooks.Open(OpenFile)
ElseIf (LCase(Action)=LCase("/Help")) Or (LCase(Action)=LCase("/?")) Then
	ShowHelp "Help"
Else
	ShowHelp "Action Not Defiend"
End If

If LCase(Action)=LCase("/Export") Then
	i = 1
	IntOURow = 2
	intSheet = 2
	objExcel.Visible = True
	Do Until objOUExcel.Cells(intSheet,1).Value = ""
		intSheet = intSheet + 1
	Loop
	objExcel.Workbooks.Add
	
	For W = 0 to intSheet - 2
		objExcel.WorkSheets.Add
	Next
	
	Do Until objOUExcel.Cells(IntOURow,1).Value = ""
		MainOU = objOUExcel.Cells(IntOURow,1)
		ExportADObjects MainOU,i
		i = i + 1
		IntOURow = IntOURow + 1
	Loop
	Set objWorkbook = objExcel.ActiveWorkbook
	objWorkbook.SaveAs(FileLoc & "-Export.xls")
ElseIf LCase(Action)=LCase("/Import") Then
	ImportADObjects
End If
ErrorMSG
Cleanup
MsgBox  "Script Completed",vbInformation + vbOKOnly, "Import/Export AD Objects"
WScript.Quit
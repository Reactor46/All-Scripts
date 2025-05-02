' ListAllUsersInGroup.vbs
'=*=*=*=*=*=*=*=*=*=*=*=*=
' Created by Assaf Miron
' Date : 14/1/07
' Modified : 11/06/08
' Description : Runs on Active Directory's Groups and lists in an Excel file 
'               Exports the Users Name,Street,Exchange Server Name and Users UPN
'=*=*=*=*=*=*=*=*=*=*=*=*=


Const ADS_SCOPE_SUBTREE = 2
Const ADS_UF_SMARTCARD_REQUIRED = &h40000  
Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000 
Const ADS_UF_ACCOUNTDISABLE = &h00002
Const GroupFileLoc = "C:\Group-Report.xls"
Const ForReading = 1
Const ForWriting = 2
Const xlCenter = -4108
Const xlSolid = 1

Set WshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject") 

Function FindObject(strObj,ObjClass)
	
	Dim objRootDSE,objConnection,objCommand,objRecordSet
	Dim strDomainLdap

	Set objRootDSE = GetObject ("LDAP://rootDSE")
	strDomainLdap  = objRootDSE.Get("defaultNamingContext")
	
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand = CreateObject("ADODB.Command")
	
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	
	Set objCommand.ActiveConnection = objConnection
	
	objCommand.CommandText = _
		"SELECT AdsPath FROM 'LDAP://" & strDomainLdap & "' WHERE objectClass='" & ObjClass & "' and sAMAccountName='" &_
			strObj & "'"
	
	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Timeout") = 30
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
	objCommand.Properties("Cache Results") = False
	
	Set objRecordSet = objCommand.Execute
	
	If objRecordSet.RecordCount = 0 Then 
		
		FindObject= 0
		
	Else 
	
		objRecordSet.Requery
		objRecordSet.MoveFirst
		Do Until objRecordSet.EOF
			FindObject= objRecordSet.Fields("AdsPath").Value
			objRecordSet.MoveNext
		Loop

	End If 	
End Function 

Sub GroupMembersInfo(GroupName)
On Error Resume Next
Set GrpMembers = GetObject(FindObject(GroupName,"group"))
objexcel.cells(introw,1) = GrpMembers.Name
objexcel.cells(introw,2) = GrpMembers.Description
For Each User in GrpMembers.Members
	If (User.Class = "user") And (User.Class <> "group") Then
		intUAC = User.Get("userAccountControl")
		If intUAC And ADS_UF_ACCOUNTDISABLE Then
'				Wscript.Echo "User Disabled (" & User.sAMAccountName & ")"
			UserDisabled = 1
		Else
			UserDisabled = 0
		End If
		If UserDisabled = 0 Then
			objexcel.cells(introw,3) = User.sAMAccountName
			objexcel.cells(introw,4)= User.givenname & " " & User.sn
			objExcel.Cells(introw,5)= User.st
			objExcel.Cells(introw,6) = User.mail
			strMailServer = User.msExchHomeServerName
			arrMailServer = Split(strMailServer,"=")
			LastCell = UBound(arrMailServer)
			If strMailServer = "" Then
				objexcel.cells(introw,7) = "No Mailbox"
			Else
			        objexcel.cells(introw,7)= arrMailServer(LastCell)
			End If
			objExcel.Cells(introw,8) = User.userPrincipalName
			intRow = intRow + 1
		End If
	Else
		If (User.Class = "group") Then
			objexcel.cells(introw,1) = User.Name
			objexcel.cells(introw,2) = User.Description
			GrpName = Mid(User.Name,4,Len(User.Name))
			GroupMembersInfo GrpName
		End If
	End If
Next

End Sub


Sub ExcelHeaders()
Set objRange = objExcel.Range("A1","I1")
objRange.Font.Size = 12
objRange.Interior.ColorIndex=15

objexcel.cells(1,1)="Group Name"
objexcel.cells(1,2)="Group Description"
objexcel.cells(1,3)="User Name"
objexcel.cells(1,4)="Full Name"
objexcel.cells(1,5)="Street"
objexcel.cells(1,6)="SMTP"
objexcel.cells(1,7)="Exchange Server Name"
objexcel.cells(1,8)="UPN"

End Sub

Sub CenterCells()
'Auto Fit the Cells
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
	objExcel.ActiveWindow.SplitRow = 1
	objExcel.ActiveWindow.FreezePanes = True
	Set ObjRange = objExcel.ActiveCell.EntireColumn
	objRange.AutoFit()
Next
End sub


' Opening File

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


'Create an Excel File
Set objExcel = CreateObject("Excel.Application")
Set objGroupExcel = CreateObject("Excel.Application")
Set objWorkbook = objGroupExcel.Workbooks.Open(FileLoc)

objExcel.Visible = True
objExcel.Workbooks.Add


i = 1
intRow = 2
IntGroupRow = 2
intSheet = 2
Do Until objGroupExcel.Cells(intSheet,1).Value = ""
	intSheet = intSheet + 1
Loop

For W = 0 to intSheet - 2
	objExcel.WorkSheets.Add
Next
Do Until objGroupExcel.Cells(IntGroupRow,1).Value = ""

	Group = objGroupExcel.Cells(IntGroupRow,1)
	Set objGroup = GetObject(FindObject(Group,"group"))
	set ObjWorkSheet=ObjExcel.WorkSheets(i)
	ObjWorkSheet.Activate
	If Len(objGroup.Description) > 31 Then
		objWorkSheet.Name = Mid(objGroup.Description,Len(objGroup.Description)/2,Len(objGroup.Description))
	Else
		objWorkSheet.Name = objGroup.Description
	End If
	'Create Nice Headers
	ExcelHeaders
	For Each Grp in objGroup.Members
		If (Grp.Class = "group") AND (Grp.Class <> "user") Then
			objexcel.cells(introw,1) = Grp.Name
			objexcel.cells(introw,2) = Grp.Description
			GroupName = Mid(Grp.Name,4,Len(Grp.Name))
			GroupMembersInfo GroupName 
		End If
		If (Grp.Class = "user") Then
			GroupName = Mid(objGroup.Name,4,Len(objGroup.Name))
			GroupMembersInfo GroupName
			Exit For
		End If
	Next

	IntGroupRow = IntGroupRow + 1
	IntRow = 2
	i = i + 1
Loop


Set objWorkbook = objExcel.ActiveWorkbook
objWorkbook.SaveAs(GroupFileLoc)
objExcel.Quit
objGroupExcel.Quit


Wscript.Echo "OK !" & vbCRLF & "Youre file has been saved in " & vbCRLF & GroupFileLoc
Wscript.Quit

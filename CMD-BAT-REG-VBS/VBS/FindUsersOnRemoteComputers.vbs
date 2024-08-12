'=*=*=*=*=*=*=*=*=*=*=*=*=
' Created by Assaf Miron
' Date : 26/11/06
' FindUsersOnRemoteComps.vbs
'=*=*=*=*=*=*=*=*=*=*=*=*=

Const ForReading = 1
Const ADS_UF_SMARTCARD_REQUIRED = &h40000 
Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000
Const ADS_SCOPE_SUBTREE = 2

On Error Resume Next

Function FindUser(strUser)
'Find a User in AD	
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
		"SELECT AdsPath FROM 'LDAP://" & strDomainLdap & "' WHERE objectClass='user' and Name='" &_
			strUser & "'"
	
	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Timeout") = 30
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
	objCommand.Properties("Cache Results") = False

	Set objRecordSet = objCommand.Execute
	
	If objRecordSet.RecordCount = 0 Then 
		
		findUser = 0
		
	Else 
		objRecordSet.Requery
		objRecordSet.MoveFirst
		Do Until objRecordSet.EOF
			findUser = objRecordSet.Fields("AdsPath").Value
			objRecordSet.MoveNext
		Loop

	End If 	
End Function 

Function IsAlive(strComputer)
Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
        ExecQuery("select * from Win32_PingStatus where address = '"_
            & strComputer& "'")
    For Each objStatus in objPing
        If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then 
            IsAlive = "machine " & strComputer& " is not reachable"
	    Exit Function
	Else
	    IsAlive = "Alive"
	    Exit Function
        End If
    Next
End Function

Sub ExcelHeaders()
'Create Excel Headers and color them in gray

Set objRange = objExcel.Range("A1","G1")
objRange.Font.Size = 12
objRange.Interior.ColorIndex=15

objExcel.cells(1,1)="User"
objExcel.cells(1,2)="First Name"
objExcel.cells(1,3)="Last Name"
objExcel.cells(1,4)="Display Name"
objExcel.cells(1,5)="Computer Name"
objExcel.cells(1,6)="Computer State"

End Sub

Sub FindCompUser(strComputer)
User = ""
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
Set colComputer = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
For Each objComputer in colComputer
    User=objComputer.UserName
Next
User = Mid(User,4,intUser)
User = FindUser(User)
If Not User = 0 Then
	Set objMember = GetObject(User)
	strUser = objMember.Name
	intUser = Len(strUser)
	objExcel.cells(introw,1)= Mid(strUser,4,intUser)
	objExcel.cells(introw,2)= objMember.givenname
	objExcel.cells(introw,3)= objMember.sn
	objExcel.cells(introw,4)= objMember.displayName
Else
	objExcel.cells(introw,1)="No User"
End If
End Sub

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

Set objExcel = CreateObject("Excel.Application")
Set objUsersExcel = CreateObject("Excel.Application")
Set objWorkbook = objUsersExcel.Workbooks.Open(FileLoc)

objExcel.Visible = True
objExcel.Workbooks.Add

introw=2
ExcelHeaders


Do Until objUsersExcel.Cells(intRow,1).Value = ""

	Comp = objUsersExcel.Cells(introw,1)
	objExcel.cells(introw,6)=IsAlive(Comp)
	If IsAlive(Comp) = "Alive" Then
		FindCompUser(Comp)
	End If
    objExcel.cells(introw,5)=Comp
	introw=introw+1

Loop

'Auto Fit the Cells
set objRange = objExcel.Range("A1")
objRange.Activate
Set ObjRange = objExcel.ActiveCell.EntireColumn
objRange.AutoFit()
set objRange = objExcel.Range("B1")
objRange.Activate
Set ObjRange = objExcel.ActiveCell.EntireColumn
objRange.AutoFit()
set objRange = objExcel.Range("C1")
objRange.Activate
Set ObjRange = objExcel.ActiveCell.EntireColumn
objRange.AutoFit()
set objRange = objExcel.Range("D1")
objRange.Activate
Set ObjRange = objExcel.ActiveCell.EntireColumn
objRange.AutoFit()
set objRange = objExcel.Range("E1")
objRange.Activate
Set ObjRange = objExcel.ActiveCell.EntireColumn
objRange.AutoFit()
set objRange = objExcel.Range("F1")
objRange.Activate
Set ObjRange = objExcel.ActiveCell.EntireColumn
objRange.AutoFit()
set objRange = objExcel.Range("G1")
objRange.Activate
Set ObjRange = objExcel.ActiveCell.EntireColumn
objRange.AutoFit()

Set objWorkbook = objExcel.ActiveWorkbook
objWorkbook.SaveAs("C:\RemoteCompUsers.xls")
objExcel.Quit

objUsersExcel.Quit


wscript.Echo "OK" & vbClrf & "The File is Saved in C:\RemoteCompUsers.Xls"
    
    

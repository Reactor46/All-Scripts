'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
' Created by Assaf Miron
' Http://assaf.miron.googlepages.com
' Date : 06/07/2009
' AddUsers-FromExcel-ToGroup.vbs
'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=

Const LogFile = "c:\UsersToGroup-LogFile.txt"
Const ForReading = 1 
Const ForWriting = 8
Const ADS_SCOPE_SUBTREE = 2
Const ADS_UF_ACCOUNTDISABLE = &h00002


Set objDialog = CreateObject("UserAccounts.CommonDialog")
Set objOUExcel = CreateObject("Excel.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
'==========
' Functions
'==========
Function FindObject(strObj,ObjClass)
' Finds an Object in the Directory and returns its ADSPath     
     Dim objRootDSE,objConnection,objCommand,objRecordSet
     Dim strDomainLdap

     Set objRootDSE = GetObject ("LDAP://rootDSE")
     strDomainLdap = objRootDSE.Get("defaultNamingContext")
     
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
          FindObject= 0 ' No Results
     Else   
          objRecordSet.Requery
          objRecordSet.MoveFirst
          Do Until objRecordSet.EOF
               FindObject= objRecordSet.Fields("AdsPath").Value
               objRecordSet.MoveNext
          Loop
     End If      
End Function
'==========
' Subs
'==========
Sub Add2Group(strUser,strGroup)
' Adding a User to a Group
	Dim objUser,objGroup
	set objGroup=GetObject(FindObject(strGroup,"group")) ' Get the Group Object
	set objUser=GetObject(FindObject(strUser,"user")) ' Get the User Object
	objGroup.Add(objUser.ADsPath) ' Add User to Group
	
	' Output success or error.
	If Err.Number <> vbEmpty Then
	    objFile.WriteLine "Failed to add The User: " & objUser.cn & " to " & objGroup.name & "."
	Else
	    objFile.WriteLine "The User: " & objUser.cn & " was added to group " & objGroup.name & ""
	End If
End Sub

'==========
' Main Code
'==========
' Opening Users Excel File
objDialog.Filter = "Excel Files|*.xls|CSV Files|*.csv"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\"
intResult = objDialog.ShowOpen
 
If intResult = 0 Then
    Wscript.Quit
Else
    FileLoc = objDialog.FileName
End If

Set objFile = objFSO.CreateTextFile(LogFile,ForWriting) ' Setting a Log File
Set objWorkbook = objOUExcel.Workbooks.Open(FileLoc) ' Opening the Excel File

objFile.WriteLine
objFile.WriteLine "Log Started Now : " & NOW
objFile.WriteLine "Opening File : " & FileLoc
objFile.WriteLine

intRow = 2
' Excel Headers Should look like this:
' User Name,Group Name
Do Until objOUExcel.Cells(intRow,1).Value = "" 'Loop Until Line is Empty
	' Get User Name from Excel
	UserID = objOUExcel.Cells(intRow,1).Value 
	' Get Group Name from Excel
	GroupToAdd = objOUExcel.Cells(intRow,2).Value
	' Add User to Group
	Add2Group UserID,GroupToAdd
	' Increment Row
	intRow = intRow + 1
Loop

objFile.WriteLine
objFile.WriteLine "Log Ended Now : " & NOW
objFile.Close
WScript.Echo "Done !" & vbNewLine & "Log File: " & LogFile
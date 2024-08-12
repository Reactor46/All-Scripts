'=*=*=*=*=*=*=*=*=*=*=*=
' Coded By Assaf Miron 
' http://assaf.miron.googlepages.com
' Description : Copy Users Between 2 Groups
'=*=*=*=*=*=*=*=*=*=*=*=

'On Error Resume Next

Const ADS_PROPERTY_APPEND = 3
Const ADS_SCOPE_SUBTREE = 2
const ForReading=1
const ForWriting=2
const ForAppending=8

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'==========
'Functions
'==========

Function FindObject(strObj,ObjClass)
     
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

'=====
'Subs
'=====

Sub getUsers(group,UF)
Dim strLDAP, objGroup, objMember

strLdap = FindObject(Group,"group") 'Getting the LDAP AdsPath of the group in the local Domain. (from RootDSE)
Set objGroup = GetObject(strLdap) 'Getting the LDAP Object to be send to the isMember function

' Write each user's DN in a Text File
For each objMember in objGroup.Members
    UF.writeline objMember.sAMAccountName
Next

UF.close

End Sub


Sub getgroups()

set objExcel=createobject("excel.application")
objexcel.workbooks.open("C:\groups.xls")
intRow=1

 do while objexcel.cells(introw,1).value <> ""

 	set UF = objFSO.OpenTextFile("c:\MoveGroups\UserFile.txt", ForWriting, 1)
	OldGroup=objexcel.cells(introw,1).value
	NewGroup=objexcel.cells(introw,2).value
	getusers OldGroup,UF
	users2group(NewGroup)
	introw=introw+1
	objFSO.DeleteFile("C:\MoveGroups\UserFile.txt")

 Loop

UF.Close
objexcel.workbooks.close

End Sub


sub users2group(group)

On Error Resume Next

dim user,logFile

set logfile = objFSO.OpenTextFile("c:\MoveGroups\LogFile.txt", ForAppending, True)

set UF = objFSO.OpenTextFile("c:\MoveGroups\UserFile.txt", ForReading)

Do Until UF.AtEndOfStream
    user=UF.readline

	set objGroup=GetObject(FindObject(group,"group"))
	set objUser=GetObject(FindObject(user,"user"))
	objGroup.Add(objUser.ADsPath)
	
	' Output success or error.
	If Err.Number <> vbEmpty Then
	    logfile.writeline "The User: " & objUser.cn & " Adding to " & objGroup.name & " failed."
	Else
	    logfile.writeline "The User: " & objUser.cn & " was added to group " & objGroup.name & ""
	End If
Loop

UF.Close
logfile.close
End Sub

'=================
' Code Begins Here
'=================

Set objFSO = CreateObject("Scripting.FileSystemObject")
If(objFSO.FolderExists("C:\MoveGroups") = False) Then
	Set objFolder = objFSO.CreateFolder("C:\MoveGroups")
End If

getgroups()


wscript.echo "Done Moving !"
wscript.quit
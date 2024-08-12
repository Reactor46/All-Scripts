Const ADS_SCOPE_SUBTREE = 2

strComputer = InputBox("Enter Machine Name to Search for:")

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

On Error Resume Next
objCommand.CommandText = "Select ADsPath From 'LDAP://dc=test,dc=tactful,dc=cloud' WHERE objectCategory = 'computer' " _
& "AND name='" & strcomputer & "'"
Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst

Do Until objRecordSet.EOF
strResult = objRecordSet.Fields("ADsPath").Value
objRecordSet.MoveNext
Loop

If strResult <> "" Then 
    Set objCompDesc = GetObject(strResult)
    strCompDesc = objCompDesc.description
	If InStr(strCompDesc, "Disabled") > 0 Then 
	strCompDisDesc = "Disabled (per Description): YES"
	Else
	strCompDisDesc = "Disabled (per Description): NO"
	End If

    strSlash = "\"
    strSlashA = "\"
    strSlashB = "\"
    strSlashC = "\"
    strSlashD = "\"
    strSlashE = "\"
    strSlashF = "\"
    strSlashG = "\"

    arrPath = Split(strResult,",")

    strOU = arrPath(UBound(arrPath) - 4)
    strOU = replace(strOU, "OU=", "")
    	If InStr(strOU, "CN=") > 0 Then 
	strOU = ""
	strSlash = ""
	End If
    strOUA = arrPath(UBound(arrPath) - 5)
    strOUA = replace(strOUA, "OU=", "")
    	If InStr(strOUA, "CN=") > 0 Then 
	strOUA = ""
	strSlashA = ""
	End If
    strOUB = arrPath(UBound(arrPath) - 6)
    strOUB = replace(strOUB, "OU=", "")
    	If InStr(strOUB, "CN=") > 0 Then 
	strOUB = ""
	strSlashB = ""
	End If
    strOUC = arrPath(UBound(arrPath) - 7)
    strOUC = replace(strOUC, "OU=", "")
    	If InStr(strOUA, "CN=") > 0 Then 
	strOUC = ""
	strSlashC = ""
	End If
    strOUD = arrPath(UBound(arrPath) - 8)
    strOUD = replace(strOUD, "OU=", "")
    	If InStr(strOUD, "CN=") > 0 Then 
	strOUD = ""
	strSlashD = ""
	End If
    strOUE = arrPath(UBound(arrPath) - 9)
    strOUE = replace(strOUE, "OU=", "")
	If InStr(strOUE, "CN=") > 0 Then 
	strOUE = ""
	strSlashE = ""
	End If
    strOUF = arrPath(UBound(arrPath) - 10)
    strOUF = replace(strOUF, "OU=", "")
	If InStr(strOUF, "CN=") > 0 Then 
	strOUF = ""
	strSlashF = ""
	End If
    strOUG = arrPath(UBound(arrPath) - 11)
    strOUG = replace(strOUG, "OU=", "")
	If InStr(strOUG, "CN=") > 0 Then 
	strOUG = ""
	strSlashG = ""
	End If

MsgBox "The Parent OU Structure For " & UCase(strComputer) & " is:  " & _
			 chr(10) & _
			 chr(10) & strCompDisDesc & _
			 chr(10) & _
			 chr(10) & strOU & strSlashA & strOUA & strSlashB & strOUB & strSlashC & strOUC & strSlashD & strOUD & strSlashE & strOUE & strSlashF & strOUF & strSlashG & strOUG
Else
MsgBox "The machine " & UCase(strComputer) & " was not found on the domain."
End If
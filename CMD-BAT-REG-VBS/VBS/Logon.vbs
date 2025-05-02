Option Explicit
Dim UserLog
Dim objFSO, arrFileLines() 
Dim objTextFile 
Dim LogDir, TempFile
Dim i
Dim k
Dim objShell, adoRecordset, objDC, objRootDSE
Dim oSysInfo, oUser, strDNSDomain, objDate  
Dim adoConnection, adoCommand, objRoot, strConfig, oResults, strQuery 
Dim nLast, nDate, oDCUser
Dim strBase, strFilter, strAttributes 
Dim arrstrDCs()
Const ForReading = 1
Const ForAppending = 8

Set oSysInfo = CreateObject("ADSystemInfo")
Set oUser = GetObject("LDAP://" & oSysInfo.UserName)
nLast = oUser.LastLogin
Set objRootDSE = GetObject("LDAP://RootDSE")
strConfig = objRootDSE.Get("configurationNamingContext")
strDNSDomain = objRootDSE.Get("defaultNamingContext")
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection
strBase = "<LDAP://" & strConfig & ">"
strFilter = "(objectClass=nTDSDSA)"
strAttributes = "AdsPath"
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 60
adoCommand.Properties("Cache Results") = False
Set adoRecordset = adoCommand.Execute
k = 0
Do Until adoRecordset.EOF
Set objDC = _
GetObject(GetObject(adoRecordset.Fields("AdsPath").Value).Parent)
ReDim Preserve arrstrDCs(k)
arrstrDCs(k) = objDC.DNSHostName
k = k + 1
adoRecordset.MoveNext
Loop
For k = 0 To Ubound(arrstrDCs)
strBase = "<LDAP://" & arrstrDCs(k) & "/" & strDNSDomain & ">"
strFilter = "(&(objectCategory=person)(objectClass=user))"
strAttributes = "distinguishedName,LastLogin"
strQuery = strBase & ";" & strFilter & ";" & strAttributes _
& ";subtree"
adoCommand.CommandText = strQuery
On Error Resume Next
Set nDate = oSysInfo.UserName("LastLogin").Value
adoRecordset.Close                                   
If nDate > nLast Then
nLast = nDate
End If
Set objFSO = CreateObject("Scripting.FileSystemObject")
Logdir = "C:\Documents and Settings\All Users\Login_log"
If Not objFSO.FolderExists(Logdir) then
objFSO.CreateFolder(Logdir)
End if
UserLog  = Logdir  & "\" & oUser.DistinguishedName & ".txt"
IF NOT objFSO.FileExists(UserLog) then
Set objTextFile = objFSO.OpenTextFile _
(UserLog, ForAppending, True)
objTextFile.WriteLine(vbCrLf & nLast)
End if
objTextFile.close
i = 0
Set objTextFile = objFSO.OpenTextFile(UserLog, ForReading)
Do Until objTextFile.AtEndOfStream
ReDim Preserve arrFileLines(i)
arrFileLines(i) = objTextFile.ReadLine
i = i + 1
Loop
If WScript.Arguments.Count = 14 Then
arrFileLines(1) = UBound(arrFileLines(i))
End if
objTextFile.Close
Set objTextFile = objFSO.OpenTextFile _
(UserLog, ForAppending, True)
objTextFile.WriteLine(vbCrLf & nLast)
objTextFile.Close
adoConnection.Close
Set objRootDSE = Nothing
Set adoConnection = Nothing
Set adoCommand = Nothing
Set adoRecordset = Nothing
Set objDC = Nothing
Set objDate = Nothing
Set objList = Nothing
Set objShell = Nothing
Set objFSO = Nothing
Set nLast = Nothing
Set nDate = Nothing
Next
Set objShell = CreateObject("Wscript.Shell")
objShell.Popup "Your last logon was on : " _
& arrFileLines(UBound(arrFileLines) - 2),10,"Logon Time",64

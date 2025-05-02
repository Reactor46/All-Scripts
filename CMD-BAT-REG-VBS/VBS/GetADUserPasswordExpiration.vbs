'#--------------------------------------------------------------------------------- 
'#The sample scripts are not supported under any Microsoft standard support 
'#program or service. The sample scripts are provided AS IS without warranty  
'#of any kind. Microsoft further disclaims all implied warranties including,  
'#without limitation, any implied warranties of merchantability or of fitness for 
'#a particular purpose. The entire risk arising out of the use or performance of  
'#the sample scripts and documentation remains with you. In no event shall 
'#Microsoft, its authors, or anyone else involved in the creation, production, or 
'#delivery of the scripts be liable for any damages whatsoever (including, 
'#without limitation, damages for loss of business profits, business interruption, 
'#loss of business information, or other pecuniary loss) arising out of the use 
'#of or inability to use the sample scripts or documentation, even if Microsoft 
'#has been advised of the possibility of such damages 
'#--------------------------------------------------------------------------------- 

Dim strFilePath, objFSO, objFile, ADOConnection, ADOCommand
Dim objRootDSE, strDNSDomain, strFilter, strQuery, ADORecordset
Dim strDN, objShell, regKey, reg, PassowrdExpired
Dim objDate, PasswordLastSetDate, regFlag, i
Dim strPosition,samAccountName
Dim currentDirectory

Set objFSO = CreateObject("Scripting.FileSystemObject")


On Error Resume Next
Const ForWriting = 2

'Get current folder location
currentDirectory = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))
Set objLogFile = objFSO.CreateTextFile(currentDirectory & "UserPasswordExpirationInfo.csv", ForWriting, True)
objLogFile.WriteLine "samAccountName,PassowrdExpired,PasswordLastSetDate"

'Get local time zone  from registry.
Set objShell = CreateObject("Wscript.Shell")
regKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias")
If (UCase(TypeName(regKey)) = "LONG") Then
    reg = regKey
ElseIf (UCase(TypeName(regKey)) = "VARIANT()") Then
    reg = 0
    For i = 0 To UBound(regKey)
        reg = reg + (regKey(i) * 256^i)
    Next
End If

Set ADOConnection = CreateObject("ADODB.Connection")
Set ADOCommand = CreateObject("ADODB.Command")
ADOConnection.Provider = "ADsDSOOBject"
ADOConnection.Open "Active Directory Provider"
Set ADOCommand.ActiveConnection = ADOConnection

' Determine the DNS domain from the RootDSE object.
Set objRootDSE = GetObject("LDAP://RootDSE")
strDomainDN = objRootDSE.Get("DefaultNamingContext")
strFilter = "(&(objectCategory=person)(objectClass=user))"


ADOCommand.CommandText = "<LDAP://" & strDomainDN & ">;" & strFilter & ";distinguishedName,pwdLastSet,userAccountControl;subtree"
ADOCommand.Properties("Page Size") = 1000

Const ADS_UF_PASSWD_CANT_CHANGE = &H40
Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000

Set ADORecordset = ADOCommand.Execute
Do Until ADORecordset.EOF
    strDN = ADORecordset.Fields("distinguishedName").Value
    regFlag = ADORecordset.Fields("userAccountControl").Value
    PassowrdExpired = True
    
    If ((regFlag And ADS_UF_PASSWD_CANT_CHANGE) <> 0) Then
        PassowrdExpired = False
    End If
    If ((regFlag And ADS_UF_DONT_EXPIRE_PASSWD) <> 0) Then
        PassowrdExpired = False
    End If

    If (TypeName(ADORecordset.Fields("pwdLastSet").Value) = "Object") Then
        Set objDate = ADORecordset.Fields("pwdLastSet").Value
        PasswordLastSetDate = ChangeTimeZone(objDate, reg)
    Else
        PasswordLastSetDate = #12/31/1600#
    End If
    
    strPosition = InStr(strDN,",")
    samAccountName = Mid(strDN,4,strPosition-4)
    
    objLogFile.WriteLine samAccountName & "," & PassowrdExpired & "," & PasswordLastSetDate
    ADORecordset.MoveNext
Loop

Function ChangeTimeZone(ByVal objDate, ByVal reg)
    Dim regAdjust, regDate, regHigh, regLow
    regAdjust = reg
    regHigh = objDate.HighPart
    regLow = objdate.LowPart

    If (regLow < 0) Then
        regHigh = regHigh + 1
    End If
    If (regHigh = 0) And (regLow = 0) Then
        regAdjust = 0
    End If
    regDate = #12/31/1600# + (((regHigh * (2 ^ 32)) + regLow) / 600000000 - regAdjust) / 1440
    
    On Error Resume Next
    ChangeTimeZone = CDate(regDate)
    If (Err.Number <> 0) Then
        On Error GoTo 0
        ChangeTimeZone = #12/31/1600#
    End If
    On Error GoTo 0
End Function

'Release Object
objFile.Close
ADOConnection.Close
ADORecordset.Close

Wscript.Echo "The script has finished running."
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

Option Explicit

Dim objShell
Dim objRoot
Dim strDomain
Dim ADOConnection, ADOCommand, ADORecordset
Dim regKey
Dim reg, i, strDN, dtmDate, objDate
Dim strBase, strFilter, strAttributes, regHigh, regLow
Dim strPosition,samAccountName

'get domain from ROOtDSE object
Set objRoot = GetObject("LDAP://RootDSE")
strDomain = objRoot.Get("defaultNamingContext")

Set ADOCommand = CreateObject("ADODB.Command")
Set ADOConnection = CreateObject("ADODB.Connection")
ADOConnection.Provider = "ADsDSOObject"
ADOConnection.Open "Active Directory Provider"
ADOCommand.ActiveConnection = ADOConnection

'Set query attribute.
strBase = "<LDAP://" & strDomain & ">"
strFilter = "(&(objectCategory=person)(objectClass=user))"
strAttributes = "distinguishedName,lastLogonTimeStamp"

Const ADS_SCOPE_SUBTREE=2
ADOCommand.CommandText = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
ADOCommand.Properties("Page Size") = 1000
ADOCOmmand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
Set ADORecordset = ADOCommand.Execute

'Get local time zone from registry
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

Do Until ADORecordset.EOF

    strDN = ADORecordset.Fields("DistinguishedName").Value
    
    On Error Resume Next
    Set objDate = ADORecordset.Fields("LastLogonTimeStamp").Value
    If (Err.Number <> 0) Then
        On Error GoTo 0
        dtmDate = #1/1/1600#
    Else
        On Error GoTo 0
        regHigh = objDate.HighPart
        regLow = objDate.LowPart
        If (regLow < 0) Then
            regHigh = regHigh + 1
        End If
        If (regHigh = 0) And (regLow = 0) Then
            dtmDate = #1/1/1600#
        Else
            dtmDate = #1/1/1600# + (((regHigh * (2 ^ 32)) + regLow)/600000000 - reg)/1440
        End If
    End If
    
    strPosition = InStr(strDN,",")
    samAccountName = Mid(strDN,4,strPosition-4)
    
    If (dtmDate = #1/1/1600#) Then
        Wscript.Echo "SamAccountName: " & samAccountName & VbCrlf & "LastLogonTimeStamp: Never Logon" & VbCrLf
    Else
        Wscript.Echo "SamAccountName: " & samAccountName & VbCrlf & "LastLogonTimeStamp: " & dtmDate & VbCrlf
    End If
    ADORecordset.MoveNext
Loop

'release object
ADORecordset.Close
ADOConnection.Close
Set objRoot = Nothing
Set objShell = Nothing
'''Logon Greeting Script

ON ERROR RESUME NEXT
Dim WSHNetwork, objDomain, DomainString, UserString, UserObj
Set WSHNetwork = CreateObject("WScript.Network")
'Automatically find the domain name
Set objDomain = getObject("LDAP://rootDse")
DomainString = objDomain.Get("dnsHostName")
'Grab the user name
UserString = WSHNetwork.UserName
'Bind to the user object to get user name and check for group memberships later
Set UserObj = GetObject("WinNT://" & DomainString & "/" & UserString)

'=======================================================
' Determine the appropriate greeting for the time of day.
'=======================================================
Dim HourNow, Greeting
HourNow = Hour(Now)
If HourNow >3 And  HourNow <12 Then
       Greeting = "Good Morning "
Else
       Greeting = "Good Afternoon "
End If
'=======================================================
'Find the Users Name

Dim GreetName
GreetName = SearchGivenName(objDomain,UserString)

' Use the Microsoft Speach API (SAPI)
'=====================================
Dim oVo
Set oVo = Wscript.CreateObject("SAPI.SpVoice")
'sets the voice
Set ovo.Voice = ovo.GetVoices("Name=Microsoft Mike").Item(0)
ovo.speak Greeting & GreetName

'Modify This Function To Change Name Format
Public Function SearchGivenName(oRootDSE, ByVal vSAN)
    ' Function:     SearchGivenName
    ' Description:  Searches the Given Name for a given SamAccountName
    ' Parameters:   RootDSE, ByVal vSAN - The SamAccountName to search
    ' Returns:      First, Last or Full Name
    ' Thanks To:    Kob3 Tek-Tips FAQ:FAQ329-5688 

    Dim oConnection, oCommand, oRecordSet
    
    Set oConnection = CreateObject("ADODB.Connection")
    oConnection.Open "Provider=ADsDSOObject;"
    Set oCommand = CreateObject("ADODB.Command")
    oCommand.ActiveConnection = oConnection
    oCommand.CommandText = "<LDAP://" & oRootDSE.get("defaultNamingContext") & _
        ">;(&(objectCategory=User)(samAccountName=" & vSAN & "));givenName,sn,name;subtree"
    Set oRecordSet = oCommand.Execute
    On Error Resume Next
    'Decide which name format to return and uncomment out 
    'that line.  Default is first name.
    'Return First Name
    SearchGivenName = oRecordSet.Fields("givenName")
    'Return Last Name
    'SearchGivenName = oRecordSet.Fields("sn")
    'Return First and Last Name
    'SearchGivenName = oRecordSet.Fields("name")
    On Error GoTo 0
    oConnection.Close
    Set oRecordSet = Nothing
    Set oCommand = Nothing
    Set oConnection = Nothing
    Set oRootDSE = Nothing
End Function


'Clean Up Memory We Used
set UserObj = Nothing
set GroupObj = Nothing
set WSHNetwork = Nothing
set DomainString = Nothing

'''End Logon Script Greeting
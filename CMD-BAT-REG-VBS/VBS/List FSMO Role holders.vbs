Option Explicit
Dim WSHNetwork, objArgs, ADOconnObj, bstrADOQueryString, RootDom, RSObj
Dim FSMOobj,CompNTDS, Computer, Path, HelpText
Dim objFSO, objFile

Set WSHNetwork = CreateObject("WScript.Network")
Set objArgs = WScript.Arguments
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("FSMO.txt")

Select Case objArgs.Count
    Case 0
        Path = InputBox("Enter your DC name or the DN for your domain"&_
                        " 'DC=MYDOM,DC=COM':","Enter path",WSHNetwork.ComputerName)
    Case 1
        Select Case UCase(objArgs(0))
            Case "?"
                WScript.Echo HelpText
                WScript.Quit
            Case "/?"
                WScript.Echo HelpText
                WScript.Quit
            Case "HELP"
                WScript.Echo HelpText
                WScript.Quit
            Case Else
                Path = objArgs(0)
        End Select
    Case Else
        WScript.Echo HelpText
        WScript.Quit
End Select

Set ADOconnObj = CreateObject("ADODB.Connection")

ADOconnObj.Provider = "ADSDSOObject"
ADOconnObj.Open "ADs Provider"


'PDC FSMO
bstrADOQueryString = "<LDAP://"&Path&">;(&(objectClass=domainDNS)(fSMORoleOwner=*));adspath;subtree"
Set RootDom = GetObject("LDAP://RootDSE")
Set RSObj = ADOconnObj.Execute(bstrADOQueryString)
Set FSMOobj = GetObject(RSObj.Fields(0).Value)
Set CompNTDS = GetObject("LDAP://" & FSMOobj.fSMORoleOwner)
Set Computer = GetObject(CompNTDS.Parent)
objFile.WriteLine "The PDC FSMO is: " & Computer.dnsHostName


'Rid FSMO
bstrADOQueryString = "<LDAP://"&Path&">;(&(objectClass=rIDManager)(fSMORoleOwner=*));adspath;subtree"

Set RSObj = ADOconnObj.Execute(bstrADOQueryString)
Set FSMOobj = GetObject(RSObj.Fields(0).Value)
Set CompNTDS = GetObject("LDAP://" & FSMOobj.fSMORoleOwner)
Set Computer = GetObject(CompNTDS.Parent)
objFile.WriteLine "The RID FSMO is: " & Computer.dnsHostName


'Infrastructure FSMO
bstrADOQueryString = "<LDAP://"&Path&">;(&(objectClass=infrastructureUpdate)(fSMORoleOwner=*));adspath;subtree"

Set RSObj = ADOconnObj.Execute(bstrADOQueryString)
Set FSMOobj = GetObject(RSObj.Fields(0).Value)
Set CompNTDS = GetObject("LDAP://" & FSMOobj.fSMORoleOwner)
Set Computer = GetObject(CompNTDS.Parent)
objFile.WriteLine "The Infrastructure FSMO is: " & Computer.dnsHostName


'Schema FSMO
bstrADOQueryString = "<LDAP://"&RootDom.Get("schemaNamingContext")&_
                     ">;(&(objectClass=dMD)(fSMORoleOwner=*));adspath;subtree"

Set RSObj = ADOconnObj.Execute(bstrADOQueryString)
Set FSMOobj = GetObject(RSObj.Fields(0).Value)
Set CompNTDS = GetObject("LDAP://" & FSMOobj.fSMORoleOwner)
Set Computer = GetObject(CompNTDS.Parent)
objFile.WriteLine "The Schema FSMO is: " & Computer.dnsHostName


'Domain Naming FSMO
bstrADOQueryString = "<LDAP://"&RootDom.Get("configurationNamingContext")&_
                     ">;(&(objectClass=crossRefContainer)(fSMORoleOwner=*));adspath;subtree"

Set RSObj = ADOconnObj.Execute(bstrADOQueryString)
Set FSMOobj = GetObject(RSObj.Fields(0).Value)
Set CompNTDS = GetObject("LDAP://" & FSMOobj.fSMORoleOwner)
Set Computer = GetObject(CompNTDS.Parent)
objFile.WriteLine "The Domain Naming FSMO is: " & Computer.dnsHostName
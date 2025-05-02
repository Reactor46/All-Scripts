
' the netbios name of the server whose mailboxes you want to return
Dim cComputerName : cComputerName = wscript.Arguments(0)
Const cWMINameSpace = "root/MicrosoftExchangeV2"
Const cWMIInstance = "Exchange_Mailbox"
Dim strWinMgmts                 ' Connection string for WMI
Dim objWMIExchange              ' Exchange Namespace WMI object
Dim listExchange_Mailboxes       ' Exchange_Mailbox collection
Dim objExchange_Mailbox         ' A single Exchange_Mailbox WMI object
strWinMgmts = "winmgmts:{impersonationLevel=impersonate}!//" & _
cComputerName & "/" & cWMINameSpace
Set objWMIExchange =  GetObject (strWinMgmts)
' Verify we were able to correctly connect to the WMI namesspace on the server
If Err.Number <> 0 Then
  WScript.Echo "ERROR: Unable to connect to the WMI namespace."
  WScript.Echo err.description & " (" & Hex (err.number) & ")"
  WScript.Quit 1
End If
 ' The Resources that currently exist appear as a list of

 ' Exchange_Mailbox instances in the Exchange namespace.
Set listExchange_Mailboxes = objWMIExchange.InstancesOf (cWMIInstance)
' Were any Exchange_Mailbox Instances returned?
If (listExchange_Mailboxes.count <= 0) Then
  set objWMIExchange = Nothing
  WScript.Echo "WARNING: No Exchange_Mailbox instances were returned."
  WScript.Quit 1
End If
' Iterate through the list of Exchange_Mailbox objects.
For Each objExchange_Mailbox in listExchange_Mailboxes

Dim strTmp, strOut
strTmp = objExchange_Mailbox.LegacyDN

set com = createobject("ADODB.Command")
set conn = createobject("ADODB.Connection")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
Com.ActiveConnection = Conn
Ldapfilter = "(&(&(&(& (mailnickname=*)(objectCategory=person)(objectClass=user)(legacyExchangeDN=" & objExchange_Mailbox.LegacyDN & ")))))"
strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & Ldapfilter & ";distinguishedName;subtree"
com.Properties("Page Size") = 100
Com.CommandText = strQuery
Set Rs1 = Com.Execute
while not Rs1.eof
	set objuser = getobject("LDAP://" & replace(rs1.fields("distinguishedName"),"/","\/"))
	set objOu = getobject(objuser.parent)
	wscript.echo "Given Name : " & objuser.givenName
	wscript.echo "SurName : " & objuser.sn
	wscript.echo "OU: " & objOu.Name
	wscript.echo "Storage Group : " & objExchange_Mailbox.StorageGroupName
	wscript.echo "Mail Store : " & objExchange_Mailbox.StoreName
	rs1.movenext
wend
strTmp = Right (strTmp, Len (strTmp) - InStrRev (strTmp, "CN="))
strTmp = Mid (strTmp, 3)
strOut = objExchange_Mailbox.MailboxDisplayName & "|" & strTmp & "|"&cComputername &"|" 
strTmp = FormatNumber (objExchange_Mailbox.Size, 0)
If Len (strTmp) < 14 Then
  strTmp = Space (14 - Len (strTmp)) & strTmp
End If
strOut = strOut & strTmp
wScript.Echo strOut

Next
set objWMIExchange = Nothing
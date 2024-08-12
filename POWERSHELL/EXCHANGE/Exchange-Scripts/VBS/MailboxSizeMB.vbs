'This script logs on to a server that is running Exchange Server and
'displays the current number of bytes that are used in the user's
'mailbox and the number of messages.

' USAGE: cscript MailboxSize.vbs SERVERNAME MAILBOXNAME

' This requires that CDO 1.21 is installed on the computer.
' This script is provided AS IS. It is intended as a SAMPLE only.
' Microsoft offers no warranty or support for this script.  
' Use at your own risk.

' Get command line arguments.
Dim obArgs
Dim cArgs

Set obArgs = WScript.Arguments
cArgs = obArgs.Count

Main

Sub Main()
   Dim oSession
   Dim oInfoStores
   Dim oInfoStore
   Dim StorageUsed
   Dim NumMessages
   Dim strProfileInfo
   Dim sMsg

   On Error Resume Next

   If cArgs <> 2 Then
      WScript.Echo "Usage: cscript MailboxSize.vbs SERVERNAME MAILBOXNAME"
      Exit Sub
   End If

   'Create Session object.
   Set oSession = CreateObject("MAPI.Session")
   if Err.Number <> 0 Then
      sMsg = "Error creating MAPI.Session."
      sMsg = sMsg & "Make sure CDO 1.21 is installed. "
      sMsg = sMsg & Err.Number & " " & Err.Description
      WScript.Echo sMsg
      Exit Sub
   End If
    
   strProfileInfo = obArgs.Item(0) & vbLf & obArgs.Item(1)

   'Log on.
   oSession.Logon , , False, True, , True, strProfileInfo
   if Err.Number <> 0 Then
      sMsg = "Error logging on: "
      sMsg = sMsg & Err.Number & " " & Err.Description
      WScript.Echo sMsg
      WScript.Echo "Server: " & obArgs.Item(0)
      WScript.Echo "Mailbox: " & obArgs.Item(1)
      Set oSession = Nothing
      Exit Sub
   End If

   'Grab the information stores.
   Set oInfoStores = oSession.InfoStores
   if Err.Number <> 0 Then

      sMsg = "Error retrieving InfoStores Collection: "
      sMsg = sMsg & Err.Number & " " & Err.Description
      WScript.Echo sMsg
      WScript.Echo "Server: " & obArgs.Item(0)
      WScript.Echo "Mailbox: " & obArgs.Item(1)
      Set oInfoStores = Nothing
      Set oSession = Nothing
      Exit Sub
   End If
    
   'Loop through information stores to find the user's mailbox.
   For Each oInfoStore In oInfoStores
      If InStr(1, oInfoStore.Name, "Mailbox - ", 1) <> 0 Then
         '&HE080003 = PR_MESSAGE_SIZE
         StorageUsed = oInfoStore.Fields(&HE080003)
         if Err.Number <> 0 Then
            sMsg = "Error retrieving PR_MESSAGE_SIZE: "
            sMsg = sMsg & Err.Number & " " & Err.Description
            WScript.Echo sMsg
            WScript.Echo "Server: " & obArgs.Item(0)
            WScript.Echo "Mailbox: " & obArgs.Item(1)
            Set oInfoStore = Nothing
            Set oInfoStores = Nothing
            Set oSession = Nothing
            Exit Sub
         End If
         
         '&H33020003 = PR_CONTENT_COUNT
         NumMessages = oInfoStore.Fields(&H36020003)

         if Err.Number <> 0 Then

            sMsg = "Error Retrieving PR_CONTENT_COUNT: "
            sMsg = sMsg & Err.Number & " " & Err.Description
            WScript.Echo sMsg
            WScript.Echo "Server: " & obArgs.Item(0)
            WScript.Echo "Mailbox: " & obArgs.Item(1)
            Set oInfoStore = Nothing
            Set oInfoStores = Nothing
            Set oSession = Nothing
            Exit Sub
         End If
         getAdinfo(obArgs.Item(1))
         sMsg = "Storage Used in " & oInfoStore.Name
         sMsg = sMsg & " (MB): " & (StorageUsed/1024)/1024
         WScript.Echo sMsg
         WScript.Echo "Number of Messages: " & NumMessages
      End If
   Next

   ' Log off.
   oSession.Logoff

   ' Clean up memory.
   Set oInfoStore = Nothing
   Set oInfoStores = Nothing
   Set oSession = Nothing 
End Sub

sub getAdinfo(usUser)

set com = createobject("ADODB.Command")
set conn = createobject("ADODB.Connection")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
Com.ActiveConnection = Conn
Ldapfilter = "(&(&(&(& (mailnickname=*)(objectCategory=person)(objectClass=user)(samaccountname=" & usUser & ")))))"
strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & Ldapfilter & ";distinguishedName;subtree"
com.Properties("Page Size") = 100
Com.CommandText = strQuery
Set Rs1 = Com.Execute
while not Rs1.eof
	set objuser = getobject("LDAP://" & replace(rs1.fields("distinguishedName"),"/","\/"))
	set objOu = getobject(objuser.parent)
	set msStore = getobject("LDAP://" +  replace(objuser.homemdb,"/","\/"))
	set soServer = getobject("LDAP://" + replace(msStore.msExchOwningServer,"/","\/"))
	set soStorageGroup = getobject(msStore.parent)
	wscript.echo "Given Name : " & objuser.givenName
	wscript.echo "SurName : " & objuser.sn
	wscript.echo "OU: " & objOu.Name
	wscript.echo "Storage Group : " & soStorageGroup.cn
	wscript.echo "Mail Store : " & msStore.cn
	rs1.movenext
wend

end sub

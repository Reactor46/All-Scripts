servername = wscript.arguments(0)
PR_HAS_RULES = &H663A000B
PR_URL_NAME = &H6707001E
PR_CREATOR = &H3FF8001E
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\MeetingDelgatesForwards.csv",2,true)
wfile.writeline("Mailbox,ForwadingAddress,Status,OU")
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
svcQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchExchangeServer)(cn=" & Servername & "));cn,name,legacyExchangeDN;subtree"
Com.ActiveConnection = Conn
Com.CommandText = svcQuery
Set Rs = Com.Execute
while not rs.eof	
	GALQueryFilter =  "(&(&(&(& (mailnickname=*)(!msExchHideFromAddressLists=TRUE)(| (&(objectCategory=person)(objectClass=user)(msExchHomeServerName=" & rs.fields("legacyExchangeDN") & ")) )))))"
	strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,mailnickname,mail;subtree"
	com.Properties("Page Size") = 100
	Com.CommandText = strQuery
	Set Rs1 = Com.Execute
	while not Rs1.eof
		call procmailboxes(servername,rs1.fields("mail"))
		wscript.echo rs1.fields("mail")
		rs1.movenext
	wend
	rs.movenext
wend
rs.close
wfile.close
set fso = nothing
set conn = nothing
set com = nothing
wscript.echo "Done"




sub procmailboxes(servername,MailboxAlias)

Set msMapiSession = CreateObject("Redemption.RDOSession")
on error Resume Next
msMapiSession.LogonExchangeMailbox MailboxAlias, servername
Set mrMailboxRules = msMapiSession.Stores.DefaultStore.Rules
if err.number = 0 then
	on error goto 0
	wscript.echo "Checking For any Delegate forwarding Rules"
	nfNonefound = 0
	for Each roRule in mrMailboxRules
		agrstr = ""
		acActType = ""
		rname = ""
		set actions = roRule.Actions 
		for i = 1 to actions.count
			acActType = actions(i).ActionType
			if acActType = 8 Then	
				nfNonefound = 1
				Wscript.echo "Delegate Rule found Forwards to"
				for each aoAdressObject In actions(i).Recipients			
					wscript.echo "Address = " & aoAdressObject.Name
					varry = verifyaddress(aoAdressObject.Address)
					vadrretval = int(varry(0))
					if vadrretval = 1 then
						wfile.writeline(mailboxAlias & "," & aoAdressObject.Address & ",Account Valid" & "," & varry(1))
						wscript.echo "Account okay"
					else
						wscript.echo "Account not valid"
						wfile.writeline(mailboxAlias & "," & aoAdressObject.Address & ",Account Invalid" & "," & varry(1))
					end if
				next	
			end If
		next
	Next
	if nfNonefound = 0 then
		wscript.echo "No Delegate forwarding rules found"
		wfile.writeline(mailboxAlias & "," & "No Delegate forwarding rules found")
	end if
else
	Wscript.echo "Error Opening Mailbox"
	wfile.writeline(mailboxAlias & "," & "Error Opening Mailbox")
end If
msMapiSession.logoff
Set msMapiSession = Nothing
Set mrMailboxRules = Nothing

End Sub

function verifyaddress(exlegancydn)

vfQuery = "<LDAP://" & strDefaultNamingContext & ">;(legacyExchangeDN=" & exlegancydn & ");name,distinguishedName;subtree"
Com.CommandText = vfQuery
ReDim retarry(2)
Set Rschk = Com.Execute
aoAccountokay = 0
While Not Rschk.EOF
	 set objUser = getobject("LDAP://" & replace(rschk.fields("distinguishedName"),"/","\/"))
	 set objOu = getobject(objuser.parent)
	 retarry(1) = objOu.distinguishedName
	 if objUser.AccountDisabled then 
		aoAccountokay = 0
         else
		aoAccountokay = 1	
	 end if
	 rschk.movenext
wend
rschk.close
set rschk = nothing
set connchk = nothing
set comchk = Nothing
retarry(0) = aoAccountokay
verifyaddress = retarry

end function
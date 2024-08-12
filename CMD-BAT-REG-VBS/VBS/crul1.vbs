Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\MeetingDelgatesForwards.csv",8,true)

PR_HAS_RULES = &H663A000B
PR_URL_NAME = &H6707001E
PR_CREATOR = &H3FF8001E
Servername = wscript.arguments(0)
MailboxAlias = wscript.arguments(1)

Set msMapiSession = CreateObject("MAPI.Session")
on error Resume next
msMapiSession.Logon "","",False,True,True,True,Servername & vbLF & MailboxAlias
if err.number = 0 then
	on error goto 0
	Set mrMailboxRules = CreateObject("MSExchange.Rules")
	mrMailboxRules.Folder = msMapiSession.Inbox	
	Wscript.echo "Checking For any Delegate forwarding Rules"
	nfNonefound = 0
	for Each roRule in mrMailboxRules
		for each aoAction in roRule.actions
			if aoAction.ActionType = 7 then
				nfNonefound = 1
				Wscript.echo "Delegate Rule found Forwards to"
				for each aoAdressObject in aoAction.arg
					Set objAddrEntry = msMapiSession.GetAddressEntry(aoAdressobject) 
					wscript.echo "Address = "  & objAddrEntry.Address
					if verifyaddress(objAddrEntry.Address) = 1 then
						wfile.writeline(mailboxAlias & "," & objAddrEntry.Address & ",Account Valid")
						wscript.echo "Account okay"
					else
						wscript.echo "Account not valid"
						wfile.writeline(mailboxAlias & "," & objAddrEntry.Address & ",Account Invalid")
					end if
				next	
			end if
		next
	next
	if nfNonefound = 0 then
		wscript.echo "No Delegate forwarding rules found"
		wfile.writeline(mailboxAlias & "," & "No Delegate forwarding rules found")
	end if
else
	Wscript.echo "Error Opening Mailbox"
	wfile.writeline(mailboxAlias & "," & "Error Opening Mailbox")
end if
Set msMapiSession = Nothing
Set mrMailboxRules = Nothing


function verifyaddress(exlegancydn)

vfQuery = "<LDAP://" & strDefaultNamingContext & ">;(legacyExchangeDN=" & exlegancydn & ");name,distinguishedName;subtree"
Com.CommandText = vfQuery
Set Rschk = Com.Execute
aoAccountokay = 0
While Not Rschk.EOF
	 set objUser = getobject("LDAP://" & replace(rschk.fields("distinguishedName"),"/","\/"))
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
set comchk = nothing
verifyaddress = aoAccountokay

end function
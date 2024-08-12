servername = wscript.arguments(0)
PR_HAS_RULES = &H663A000B
PR_URL_NAME = &H6707001E
PR_CREATOR = &H3FF8001E
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\mbxforwardingRules.csv",2,true)
wfile.writeline("Mailbox,FolderPath,Creator,AdressObject,SMTPForwdingAddress")
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
	strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,mailnickname;subtree"
	com.Properties("Page Size") = 100
	Com.CommandText = strQuery
	Set Rs1 = Com.Execute
	while not Rs1.eof
		call procmailboxes(servername,rs1.fields("mailnickname"))
		wscript.echo rs1.fields("mailnickname")
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


Sub procmailboxes(servername,mailboxname)

set objSession = CreateObject("MAPI.Session") 
objSession.Logon "","",false,true,true,true,servername & vbLF & mailboxname
Set objInfoStores = objSession.InfoStores 
set objInfoStore = objSession.GetInfoStore 
set Inbox = objSession.Inbox
if inbox.fields.item(PR_HAS_RULES) = true then
	isforward = false
	Set objMessages = inbox.HiddenMessages
	for Each objMessage in objMessages 
		if objMessage.type = "IPM.Rule.Message" then
 			isforward = procrule(objMessage,mailboxname,inbox.fields.item(PR_URL_NAME).value)
			if isforward = true then
				objMessage.delete
			end if 
		end if
	next
	if isforward = true then
	for Each objMessage in objMessages 
		if objMessage.type = "IPM.RuleOrganizer" then
			Wscript.echo "Deleteing Rule"
			objMessage.delete
		end if
	next
	end if
end if 

end sub

function procrule(objmessage,MailboxName,folderpath)
frule = false
splitarry = split(hextotext(objmessage.fields.item(&H65EF0102)),chr(132),-1,1)
if ubound(splitarry) <> 0 then
	wscript.echo 
    	wscript.echo "Mailbox Name :" & MailboxName
	wscript.echo "Folder Path :" & folderpath
	wscript.echo "Rule Created By : " & objmessage.fields.item(PR_CREATOR).value
	mbname = MailboxName
	fpath = folderpath
	creator = objmessage.fields.item(PR_CREATOR).value
	frule = true
end if
tfirst = 0
addcount = 1
for i = 0 to ubound(splitarry)
	addrrsplit = split(splitarry(i),chr(176),-1,1)
	for j = 0 to ubound(addrrsplit)
		addrcontsep = chr(3) & "0"
		if instr(addrrsplit(j),addrcontsep) then 
			if tfirst = 1 then addcount = addcount + 1
			wscript.echo 
			wscript.echo "Address Object :" & addcount
			redim Preserve resarray(1,1,1,1,1,addcount)
			resarray(1,0,0,0,0,addcount) = mbname
			resarray(1,1,0,0,0,addcount) = fpath
			resarray(1,1,1,0,0,addcount) = creator		
			if instr(addrrsplit(j),"0/o=") then 
				resarray(1,1,1,1,0,addcount) = mid(addrrsplit(j),(instr(addrrsplit(j),"0/o=")+1),len(addrrsplit(j)))
				WScript.echo "ExchangeDN :" & mid(addrrsplit(j),(instr(addrrsplit(j),"0/o=")+1),len(addrrsplit(j)))
			else 
				WScript.echo "Address :" & mid(addrrsplit(j),3,len(addrrsplit(j)))
				resarray(1,1,1,1,0,addcount) = mid(addrrsplit(j),3,len(addrrsplit(j)))
			end if 
			tfirst = 1		
		end if
		smtpsep = Chr(254) & "9"
		if instr(addrrsplit(j),smtpsep) then 
			slen = instr(addrrsplit(j),smtpsep) + 2
			elen = instr(addrrsplit(j),chr(3))
			Wscript.echo "SMTP Forwarding Address : " & mid(addrrsplit(j),slen,(elen-slen))
			resarray(1,1,1,1,1,addcount) = mid(addrrsplit(j),slen,(elen-slen))
		end if
	next
next
if frule = true then
	for r = 1 to ubound(resarray,6)
		wfile.writeline(resarray(1,0,0,0,0,r) & "," & resarray(1,1,0,0,0,r) & "," & resarray(1,1,1,0,0,r) & "," & resarray(1,1,1,1,0,r) & "," & resarray(1,1,1,1,1,r))
	next
	procrule = true
else
	procrule = false
end if

end function


Function hextotext(binprop)
arrnum = len(binprop)/2
redim aout(arrnum)
slen = 1
for i = 1 to arrnum
	if CLng("&H" & mid(binprop,slen,2)) <> 0 then
		aOut(i) = chr(CLng("&H" & mid(binprop,slen,2)))
		rem wscript.echo CLng("&H" & mid(binprop,slen,2)) & "," & chr(CLng("&H" & mid(binprop,slen,2)))
	end if
	slen = slen+2
next
hextotext = join(aOUt,"")
end function




set req = createobject("microsoft.xmlhttp")

DestinationServername = "servername.domain"
DestinationMailbox = "https://servername.domain/exchange/mailbox"
Destinationusername = "user"
Destinationpassword = "password"
Destinationdomain = "domain"


dstchkary = FBAAuth(DestinationServername,Destinationusername,Destinationpassword,Destinationdomain)

call CreateFolder("testNewFolder","https://servername.domain/public/testFolder12")
Sub CreateFolder(DisplayName,Href)
req.open "MKCOL", Href, false
req.setrequestheader "Content-Type", "text/xml"
req.SetRequestHeader "cookie", dstchkary(0)
req.SetRequestHeader "cookie", dstchkary(1)
req.setRequestHeader "Translate","f"
req.send strxml
if req.status = 201 then
	Wscript.echo "Folder created sucessfully"
else
	wscript.echo req.status
	wscript.echo req.statustext
end if

end Sub

Function FBAAuth(snServername,mnMailboxname,strpassword,strdomain)

strusername =  strdomain & "\" & mnMailboxname
szXml = "destination=https://" & snServername & "/exchange/&flags=0&username=" & strusername 
szXml = szXml & "&password=" & strpassword & "&SubmitCreds=Log On&forcedownlevel=0&trusted=0"
req.Open "post", "https://" & snServername & "/exchweb/bin/auth/owaauth.dll", False
req.send szXml
reqhedrarry = split(req.GetAllResponseHeaders(), vbCrLf,-1,1)
dim chkary(1)
for c = lbound(reqhedrarry) to ubound(reqhedrarry)
	if instr(lcase(reqhedrarry(c)),"set-cookie: sessionid=") then chkary(0) = right(reqhedrarry(c),len(reqhedrarry(c))-12)
	if instr(lcase(reqhedrarry(c)),"set-cookie: cadata=") then chkary(1) = right(reqhedrarry(c),len(reqhedrarry(c))-12)
next
FBAAuth = chkary

End function

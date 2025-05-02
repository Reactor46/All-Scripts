set req = createobject("microsoft.xmlhttp")

Destinationusername = "domain\user"
Destinationpassword = "password"


call CreateFolder("testNewFolder","http://servername/public/testFolder121")

Sub CreateFolder(DisplayName,Href)
req.open "MKCOL", Href, false,Destinationusername,Destinationpassword
req.setrequestheader "Content-Type", "text/xml"
req.setRequestHeader "Translate","f"
req.send strxml
if req.status = 201 then
	Wscript.echo "Folder created sucessfully"
else
	wscript.echo req.status
	wscript.echo req.statustext
end if

end Sub


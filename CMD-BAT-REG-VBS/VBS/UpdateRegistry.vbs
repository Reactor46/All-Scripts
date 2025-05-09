Dim configObj, loaded, WshShell
Set configObj = CreateObject("MSXML2.DOMDocument")


Set WshShell = WScript.CreateObject("WScript.Shell")

configObj.async = false
loaded = configObj.load(WScript.Arguments(0))

if not loaded then
	wscript.echo "Error: Failed to load the document: " & configObj.parserError.reason
	'return the error string
end if

On Error Resume Next

''Update ManagementDB registry keys
dim node : set node = configObj.selectSingleNode("/UpdateConfiguration/ManagementDB")
if not (node Is Nothing) then

	wscript.echo "Info: Updating the management database registry keys"

	newDBServer = node.getAttribute("newDBServer")
	newDBName = node.getAttribute("newDBName")

	bKey = WshShell.RegRead("HKLM\SOFTWARE\Microsoft\BizTalk Server\3.0\Administration\MgmtDBServer")
	if (err = 0) then
		WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\BizTalk Server\3.0\Administration\MgmtDBServer", newDBServer, "REG_SZ"
		WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\BizTalk Server\3.0\Administration\MgmtDBName", newDBName, "REG_SZ"

		'lets look for 64 bit registry keys
		bKey = WshShell.RegRead("HKLM\SOFTWARE\Wow6432Node\Microsoft\BizTalk Server\3.0\Administration\MgmtDBServer")
		if (err = 0) then
			WshShell.RegWrite "HKLM\SOFTWARE\Wow6432Node\Microsoft\BizTalk Server\3.0\Administration\MgmtDBServer", newDBServer, "REG_SZ"
			WshShell.RegWrite "HKLM\SOFTWARE\Wow6432Node\Microsoft\BizTalk Server\3.0\Administration\MgmtDBName", newDBName, "REG_SZ"
		else
			wscript.echo "Info: No 64 bit keys were found for the management database"
		end if
		

	else
		wscript.echo "Error: The management database registry keys were not updated because no registry keys could be found."
	end if
	

	wscript.echo "Info: Done updating the management database"
else
	wscript.echo "Info: The management database was not updated because no node was found in the xml file."

end if

''Update RuleEngineDB registry keys
err.Clear
set node = configObj.selectSingleNode("/UpdateConfiguration/RuleEngineDB")
if not (node Is Nothing) then
	
	wscript.echo "Info: Updating the RuleEngine database"

	newDBServer = node.getAttribute("newDBServer")
	newDBName = node.getAttribute("newDBName")
	
	bKey = WshShell.RegRead("HKLM\SOFTWARE\Microsoft\BusinessRules\3.0\DatabaseServer")
	if (err = 0) then
		WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\BusinessRules\3.0\DatabaseServer", newDBServer, "REG_SZ"
		WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\BusinessRules\3.0\DatabaseName", newDBName, "REG_SZ"
	else
		err.Clear
		'lets look for 64 bit registry keys
		bKey = WshShell.RegRead("HKLM\SOFTWARE\Wow6432Node\Microsoft\BusinessRules\3.0\DatabaseServer")
		if (err = 0) then
			WshShell.RegWrite "HKLM\SOFTWARE\Wow6432Node\Microsoft\BusinessRules\3.0\DatabaseServer", newDBServer, "REG_SZ"
			WshShell.RegWrite "HKLM\SOFTWARE\Wow6432Node\Microsoft\BusinessRules\3.0\DatabaseName", newDBName, "REG_SZ"
		else
			wscript.echo "Error: Rule Engine registry keys were not updated because no registry keys could be found."
		end if
	end if

	wscript.echo "Info: Done updating the RuleEngine database"

	wscript.echo "Info: Updating the RuleComposer database"

	err.Clear

	bKey = WshShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\BusinessRules\3.0\RuleComposer\RuleStoreServer")
	if (err = 0) then
		WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\BusinessRules\3.0\RuleComposer\RuleStoreServer", newDBServer, "REG_SZ"
		WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\BusinessRules\3.0\RuleComposer\RuleStoreDatabase", newDBName, "REG_SZ"
	end if

	wscript.echo "Info: Done updating the RuleComposer database"

else
	wscript.echo "Info: The Rule Engine database was not updated because no node was found in the xml file."
end if

err.Clear
'sso
set node = configObj.selectSingleNode("/UpdateConfiguration/OtherDatabases/Database[@Name='SSO']")
if not (node Is Nothing) then
	wscript.echo "Info: Updating the SSO database"

	newDBServer = node.getAttribute("newDBServer")
	newDBName = node.getAttribute("newDBName")

	WshShell.Run """%CommonProgramFiles%\Enterprise Single Sign-On\Ssoconfig.exe"" -setdb " + newDBServer + " " + newDBName, 4, True

	wscript.echo "Info: Done updating the SSO database"
else
	wscript.echo "Info: The SSO database was not updated because no node was found in the xml file."
end if




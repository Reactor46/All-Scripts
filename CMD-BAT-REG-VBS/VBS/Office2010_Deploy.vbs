dim oWMI, oItem, oShell
dim colItems
dim strVersion, strCmd

strComputer = "."
strCmd = ""
set oWMI = GetObject("winmgmts:\\" & COMPUTER & "\root\CIMV2")
set colItems = oWMI.ExecQuery("SELECT * FROM win32_Product WHERE Name LIKE 'Microsoft Office%'")
set oShell = CreateObject("WScript.shell")
	
if colItems.Count >= 1 then
	for each oItem in colItems
		select case oItem.Name
			case "Microsoft Office Standard 2007"
				strCmd = "\\server\share\setup.exe /config \\server\share\Standard.WW\config.xml" 'replace with your install share
				Exit For
			case "Microsoft Office Professional Plus 2007"
				strCmd = "\\server\share\setup.exe /config \\server\share\ProPlus.WW\config.xml"  'replace with your install share
				Exit For
		end select
	next
else
	strCmd = ""	
end if
	
if strCmd <> "" then
	oShell.run strCmd, 10 , true
end if

set oWMI = nothing
set oItem = nothing
set colItems = nothing
set oShell = nothing
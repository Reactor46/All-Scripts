Set wshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

Set fil = fso.OpenTextFile(".\aps.csv",1) ' simple text file with the following format "name,ip address,mac address"

strScope = "192.168.1.0" ' this should match the scope that you want the reservations added to

Do While Not fil.AtEndOfStream
	strTemp = fil.ReadLine
	arrInfo = Split(strTemp,",")
	strName = arrInfo(0)
	strIP = arrInfo(1)
	strMAC = arrInfo(2)
	wshShell.run "netsh dhcp server scope " & strScope & " add reservedip " & strIP & " " &  strMAC & " " & strName,,True

Loop

fil.Close
Set fil = Nothing
Set fso = Nothing
Set wshShell = Nothing	

WScript.Echo "Finished."
WScript.Quit
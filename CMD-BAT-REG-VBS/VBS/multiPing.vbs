' MultiPing.vbs
'=*=*=*=*=*=*=*=*=*=*=*=*=
' Created by Assaf Miron
' Date : 19.08.04
' Update : 30.07.05
' Description : Pings a List of Computers Defiend in a Config File
'               And Closes all the Opend CMD Windows opend          
'=*=*=*=*=*=*=*=*=*=*=*=*=


Const AppConfig = "MultiPing.Config"
Const ForReading = 1

Dim ObjConfigFile
Dim objDictionary
Dim oArgs

Set objShell = createObject("Shell.Application")
Set objFSO = CreateObject ("Scripting.FileSystemObject")
Set objDictionary = CreateObject("Scripting.Dictionary") 
Set ObjConfigFile = objFSO.OpenTextFile(AppConfig,ForReading)
Set oArgs = WScript.Arguments

' Read Config File
Do While ObjConfigFile.AtEndOfStream <> True
	strLine  = ObjConfigFile.ReadLine	
	arrLine  = Split(strLine,":")
	' Create Dictionary Table for all the Keys in the Config File
	objDictionary.Add UCase(arrLine(0)),arrLine(1)
Loop

If oArgs.Count < 1 Then
	ColKeys = objDictionary.Keys
	strMSG = "Run the Script with one of the following Arguments:" & vbNewLine
	For Each strKey In ColKeys
		strMSG = strMSG & vbTab & strKey & vbNewLine
	Next
	strMSG = strMSG & "You can Edit the " & AppConfig & " File to insert more MultiPing Servers"
	strMSG = StrMSG & vbNewLine & "First Enter an Alias for the List and then the Servers Name You want to Ping."
	strMSG = StrMSG & vbNewLine & "Domain:DC01,DC02,DNS01,DNS02,WINS"
	strMSG = StrMSG & vbNewLine & vbNewLine & "The Script will Now Exit"
	WScript.Echo strMSG
	WScript.Quit
End If

' Minimizing All Windows Before Running
objShell.MinimizeAll

' Pinging (loop) All Servers
Set WshShell = createObject("WScript.Shell")

arrServers = Split(objDictionary.Item(UCase(oArgs(0))),",")

For Each SRV In arrServers
	WshShell.run("cmd /k ping -t " & SRV)
Next

' Wait for all CMD's to open
wscript.sleep 500

' Tile All windows Horizontally
objShell.TileHorizontally
set objShell = Nothing

wscript.echo "Close ?"

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'ping.exe'")
For Each objProcess in colProcessList
    objProcess.Terminate()
Next
Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'cmd.exe'")
For Each objProcess in colProcessList
    objProcess.Terminate()
Next
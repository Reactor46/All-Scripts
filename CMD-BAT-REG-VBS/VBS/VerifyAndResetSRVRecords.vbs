'---------------------------------------------------------------------------------
' The sample scripts are not supported under any Microsoft standard support
' program or service. The sample scripts are provided AS IS without warranty
' of any kind. Microsoft further disclaims all implied warranties including,
' without limitation, any implied warranties of merchantability or of fitness for
' a particular purpose. The entire risk arising out of the use or performance of
' the sample scripts and documentation remains with you. In no event shall
' Microsoft, its authors, or anyone else involved in the creation, production, or
' delivery of the scripts be liable for any damages whatsoever (including,
' without limitation, damages for loss of business profits, business interruption,
' loss of business information, or other pecuniary loss) arising out of the use
' of or inability to use the sample scripts or documentation, even if Microsoft
' has been advised of the possibility of such damages.
'---------------------------------------------------------------------------------
Option Explicit
' ################################################
' The starting point of execution for this script.
' ################################################
Sub Main()
	Dim Input 
	Dim Wmiobject,Computernames, Computername
	Dim DR,StrComputer,RoleNumber
	StrComputer ="."
	Set Wmiobject = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" & StrComputer& "\root\cimv2")
	Set Computernames =  Wmiobject.ExecQuery _
	("Select DomainRole from Win32_ComputerSystem")
	For Each Computername In Computernames
		RoleNumber = Computername.DomainRole
	Next
	If RoleNumber = 5  Then 
		Input = inputbox("Please input 'Reset' or 'GetSrv'" )
		If UCase(input) = "RESET" Then 
			Call ResetNetLogon
		ElseIf UCase(Input) = "GETSRV" Then 
			Call GetSRVRecord
		Else 
			If Input = Empty Then 
				wscript.quit
			Else 
				wscript.echo "Please input 'Reset'or 'GetSRV'"
			End If 
		End If 
	Else 
		WScript.echo "Please run the script on a domain controller"
	End If 
End Sub 

' ################################################
'This funtion is to reset Netlogon service
' ################################################
Function ResetNetLogon()
	Dim NetlogonDNSFile,NetlogonDNBFile
	Dim Objshell,Fs,Thisday,File,ObjWMIService,StrComputer
	Dim StrService,ColListOfServices,ObjService
	StrComputer = "."
	NetlogonDNSFile="c:\Windows\System32\config\netlogon.dns"
	NetlogonDNBFile="c:\Windows\System32\config\netlogon.dnb"
	Set Objshell =createobject("WScript.shell")
	Set Fs = createobject("scripting.filesystemobject")
	Thisday = Today_Date()
	If Fs.FileExists(NetlogonDNSFile) And Fs.FileExists(NetlogonDNBFile) Then  
		Set file =Fs.getfile(NetlogonDNSFile) 
		file.name = "Oldnetlogon_"&Thisday&".dns" 'Rename the netlogon.dns file
		Set File = Nothing 
		Set File =Fs.getFile(NetlogonDNBFile) 
		File.name = "Oldnetlogon_"&Thisday&".dnb" 'Rename the netlogon.dnd File
		Set File = Nothing 
		Set ObjWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" & StrComputer& "\root\cimv2")
		StrService = " 'Netlogon' "
		Set ColListOfServices = ObjWMIService.ExecQuery _
		("Select * from Win32_Service Where Name ="_
		& StrService & " ") 'Get netlogon Service instance
		For Each ObjService in ColListOfServices
			ObjService.StopService() 'Stop netlogon service
			WScript.Echo "Your "& StrService & " service has stoped" 
			WScript.sleep 3000 
			ObjService.StartService() 'Start netlogon service
			WScript.Echo "Your "& StrService & " service has reseted" 
			WScript.Quit
		Next 
	Else 
		msgbox "The Files does not exist ,please check."
	End If 
End Function 
' ################################################
'This funtion is to get SRV record
' ################################################
Function  GetSRVRecord()
	Dim ObjFSO,NetlogonDNSfile
	Dim ObjRegEx,Objfile,Result,ObjReadfile
	Dim StrSearchString,ColMatches,StrMatch
	Set ObjFSO = createobject("scripting.Filesystemobject")
	NetlogonDNSFile="c:\Windows\System32\config\netlogon.dns"
	If ObjFSO.Fileexists(NetlogonDNSFile) Then   'Verify if the file exists
		Set ObjReadFile = objFSO.OpenTextFile(NetlogonDNSFile, 1, False)
		Const ForReading = 1
		Set ObjRegEx = CreateObject("VBScript.RegExp")
		ObjRegEx.Pattern = "IN SRV"
		Set ObjFSO = CreateObject("Scripting.FileSystemObject")
		Set ObjFile = ObjFSO.OpenTextFile(NetlogonDNSFile, ForReading)
		Result = ""
		Do Until ObjFile.AtEndOfStream	 'Get the SRV record and output them
			StrSearchString = ObjFile.ReadLine
			Set ColMatches = ObjRegEx.Execute(Ucase(strSearchString))  
			If ColMatches.Count > 0 Then
				For Each StrMatch in colMatches   
					Result =  Result & vbCRLF & StrSearchString
				Next
			End If
		Loop
		WScript.echo result
		ObjFile.Close
	Else 
		WScript.echo "The file not exists.Maybe you have reset netlogon just now ,please wait a moment!"
	End If  
End Function 

' ################################################
'This funtion gets the date and time in format yymmddhhmmss 
' ################################################

Function Today_Date()
	Today_Date=Right(Year(Date),4) & Right("0" & Month(Date),2) & Right("0" & Day(Date),2)&Right("0"&Hour(Time),2)&Right("0"&Minute(Time),2)&Right("0"&second(Time),2)
End Function
' **********

Call Main
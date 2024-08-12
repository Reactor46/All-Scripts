'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 2007
'
' NAME: 
'
' AUTHOR: Jason Johnson , Horry County Schools
' DATE  : 1/5/2009
'
' COMMENT: 
'
'==========================================================================

'***********************************
' ------ SCRIPT CONFIGURATION ------
'***********************************

strOU = "OU=NVNV,OU=FACILITIES,DC=phsi,DC=primehealthcare,DC=com" ' Change this line to point to the container you want to run the command on

strUserID = InputBox("Please enter a username:","Required Info","")
strUserPass = InputBox("Please enter your password:","Required Info","")

'Insert routine specific configuration below

'***********************************
' ------ END CONFIGURATION ---------
'***********************************

On Error Resume Next

'*****************************************
'Insert the code for the action you want to perform into the following subroutine
'*****************************************
Sub PerformAction(strCompName)

	Set wshshell = CreateObject("wscript.shell")
	If strUser <> "" And strPass <> "" then
		wshshell.Run "psexec.exe -accepteula \\" & strCompName & " -u " & strUserID & " -p " & strUserPass & " cmd /c " & chr(34) & "echo . | C:\Scripts\Batch-CMD\Hosts_Add.bat y n" & Chr(34),,False
	End If
	Set wshshell = Nothing

End Sub

'*****************************************
'The ModifyUsers routine loops through the computers
'in the specified OU and attempts to Ping them.
'If they are reachable, PerformAction routine
'is called.
'*****************************************
Sub ModifyUsers(oObject)

Dim oComp
For Each oComp in oObject
	'WScript.Echo oComp.class
	Select Case oComp.Class
		Case "computer"
			strComputer = mid(oComp.name,4)
			'If strComputer <> "phoenix" then
			Set cPingResults = GetObject("winmgmts:{impersonationLevel=impersonate}//" & _
				"localhost/root/cimv2"). ExecQuery("SELECT * FROM Win32_PingStatus " & _
				"WHERE Address = '" + strComputer + "'")

			blPingable = False
			For Each oPingResult In cPingResults
				If oPingResult.StatusCode = 0 Then
					blPingable = True
				End If
			Next
			If blPingable = True Then
				PerformAction strComputer
			Else
				WScript.Echo "Unable to contact " & strComputer
			End If 
		Case "container", "organizationalUnit"
			ModifyUsers(oComp)
	End Select
Next

End Sub


Dim oDomain
Set oDomain=GetObject("LDAP://" & strOU)

ModifyUsers(oDomain)

Set oDomain = Nothing

MsgBox "Finished"

WScript.Quit

'#============================================================================================================================================================
'#  SCRIPT.........:	IE8_Deploy.vbs
'#  AUTHOR.........:	CriscoLeet
'#  VERSION........:	3.0
'#  DATE...........:	05/1/12
'#  LICENSE........:	Freeware
'#  REQUIREMENTS...:  
'#
'#  DESCRIPTION....:	Installs Internet Explorer 8 on Windows XP 32-bit workstations.
'#
'#  NOTES..........:	Added function to check OS version and file on root, if script does not meet requirement it terminates.
'# 
'#  REFERENCES.....:	http://stackoverflow.com/questions/10098694/vbscript-if-statement-to-determine-os-version-and-servicepack
'#  REFERENCES.....:	http://www.activexperts.com/activmonitor/windowsmanagement/wmi/samples/performancecountersnet/
'#  REFERENCES.....:	http://www.activexperts.com/activmonitor/windowsmanagement/adminscripts/computermanagement/software/
'#  REFERENCES.....:	http://blogs.technet.com/b/heyscriptingguy/archive/2004/08/16/how-can-i-give-a-user-a-yes-no-prompt.aspx
'#  REFERENCES.....:   
'#============================================================================================================================================================

strComputer = "."

' Examines workstation for Operating System Version.  If workstation does not meet criteria then script terminates.
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colOperatingSystem = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

For Each objOperatingSystem in colOperatingSystem
ServicePack = objOperatingSystem.ServicePackMajorVersion
Version = objOperatingSystem.Version

Next

IF Mid(Version,1,3)="5.1" Then	

' Examines workstation for specified file.  If file is not found scipt continues execution.
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objfso.FileExists("C:\noUMCupdates.txt") = True Then
	WScript.Quit()
	End If	

		If objfso.FileExists("C:\Program Files\Internet Explorer\xpshims.dll") = False Then

' StartMessage.vbs is sent to display notification to user.
			Dim strFileName
			strFileName =  Chr(34) & "\\server\StartMessage.vbs" & Chr(34)

			Set WSHShell = CreateObject("WScript.Shell") 
			WSHShell.Run "wscript " & strFileName, , False
		
' Internet Explorer 8 begins initialization on workstation (created using Windows IEAK 8).
			Const ALL_USERS = True
			Set objService = GetObject("winmgmts:")
			Set objSoftware = objService.Get("Win32_Product")
			errReturn = objSoftware.Install("\\server\IE8-Setup-Full.msi", , ALL_USERS)
		
' Notification is sent upon completion and prompt for reboot.
			intAnswer = _
			Msgbox("Installation is complete, do you want to reboot?  A restart is required for full functionality.", _
			vbYesNo, "Information Systems Notification")

				If intAnswer = vbYes Then
					Set objShell = WScript.CreateObject("WScript.Shell")
					objShell.Run "C:\WINDOWS\system32\shutdown.exe -r -t 0"
				Else
					WScript.Quit()
				End If
	
		Else
			WScript.Quit()
		End If
		
Else
	WScript.Quit()
END IF
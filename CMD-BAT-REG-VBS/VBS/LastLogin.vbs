
Dim WshShell, strMessage, strDomain, strUserName, strIPAddress, objuser, objBadLogins, objPwdLastChanged, objLastFailedLogin

 

on error resume Next

 

' Set WshShell

Set WshShell = WScript.CreateObject("WScript.Shell")

 

' Pull Environment variables for domain/user

strDomain = WshShell.ExpandEnvironmentStrings("%USERDOMAIN%")

strUserName = WshShell.ExpandEnvironmentStrings("%USERNAME%")

strIPAddress = GetDNSAddresses()

objDomain = "Domain"       

objUserName = strUserName

objIpAddress = strIpAddress


'Attempt to bind to the user

Set objUser = GetObject("WinNT://"& objDomain &"/"& objUserName, user)

 

'How many wrong logins?

objFailedLogins = objUser.FailedPasswordAttempts

 

' Calculate the date the password was last changed

objPwdLastChanged = CStr(objUser.PasswordExpirationDate - objUser.Get("MaxPasswordAge") / (60 * 60 * 24))

 
strMessage = "User Name:		" & objUserName & vbCRLF

strMessage = strMessage & "Last Succesful Login:		" & objUser.LastLogin & vbCRLF

strMessage = strMessage & "Date Account Will Expire:		" & objUser.AccountExpirationDate

DIM fso, GPInfoFile,strPath
strPath = WshShell.ExpandEnvironmentStrings("%userprofile%")
Set fso = CreateObject("Scripting.FileSystemObject")
Set GPInfoFile= fso.CreateTextFile(strPath &"\Downloads\userinfo.txt", True)
GPInfoFile.WriteLine(strMessage) 
GPInfoFile.Close

strcomputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
    
strCount = 1

For Each objitem in colitems
    If strCount = 1 Then
        strIPAddress = Join(objitem.IPAddress, ",")
        IP = stripaddress
        strCount = strCount + 1
        'wscript.echo IP
     		Dim IPInfoFile
			Set IPInfoFile= fso.CreateTextFile(strPath &"\Downloads\IpConfigResults.txt", True)
			IPInfoFile.WriteLine(IP) 
			IPInfoFile.Close
     
     End If

next

Dim oVo

Dim usrinfo

Set oVo = Wscript.CreateObject("SAPI.SpVoice")



usrinfo = wshshell.run("bginfo2.exe userinfo.bgi /timer:0",0,true)
 
Dim strCrDate, strDateDiff, strAcctExp

Set wshShell = WScript.CreateObject("WScript.Shell")

strCrDate = Now()

strAcctExp = objUser.AccountExpirationDate

strDateDiff = DateDiff("d", strCrDate, strAcctExp) 

If strDateDiff > 0 And strDateDiff <= 45 Then
   msgbox "Warning, your account will expire in " & strDateDiff & _
             " days (" & strAcctExp & ")." & Chr(10) & Chr(10) & _
          "Please take the current Information Assurance Awareness" & Chr(10) & _
	  "and Information Protection courses located on the Air Force Portal    (ADLS, Annual Total Force Awareness)." & Chr(10) & Chr(10) & _
	  "Your expiration date WILL NOT update automatically. Your CSA will accomplish this task." & Chr(10) & _
	  "Upon course completion, please give a copy of your certificate to your CSA." & Chr(10) & _ 
	  "If your training is up to date (less than 1 year old), please" & Chr(10) & _
	  "contact the Comm Focal Point, x7980, for further guidance.", 48, "Your Account is about to expire"

Dim Greeting

Greeting = "your account will expire in " & strDateDiff & "days please reaccomplish your ia training and contact your See ess A"

oVo.speak Greeting

End If


If strUserName = "james.black.adm" OR strUserName = "dominique.hatche.adm" OR strUserName = "maria.ratliff.adm" THEN 
 
usrinfo = wshshell.run("bginfo2.exe adminuserinfo.bgi /timer:0",0,true)

 msgbox "Warning, you are logged in with an ADMINISTRATOR account" & Chr(10) & _
          "Actions made with this account Will have elevated priveleges." & Chr(10) & _
	  "Please exercise increased caution when running scripts and opening files." & Chr(10) & Chr(10) & _
	  "Any malicious logic encountered while logged on with this account will execute" & Chr(10) & _
	  "with elevated privelges and could potentially have wide reaching effects." 
End IF

'Logon Control Protocol
'If strUserName = "" Then
'Set objWMIService = GetObject("winmgmts:" _
'    & "{impersonationLevel=impersonate,(Shutdown)}!\\" & _
'        strComputer & "\root\cimv2")

'Set colOperatingSystems = objWMIService.ExecQuery _
'    ("Select * from Win32_OperatingSystem")

'For Each objOperatingSystem in colOperatingSystems
'    objOperatingSystem.Reboot()
         
'Next
 
'usrinfo2 = wshshell.run("bginfo2.exe userinfospecial.bgi /timer:0",0,true)

'wscript.Quit

'End If

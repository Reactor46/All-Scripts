'Unsolicited Remote Assistance version 1.2_RD - 9/24/10
'Thanks to Kevin Bumber
'
'Remotely control a computer via Remote Assistance without requiring the user
' to either accept the incoming connection or grant the 'Expert' remote
' control permissions.
'
'*** Description:
'
'This script allows you to connect to a computer via Remote Assistance without
' requiring remote assistance solicitation from the user needing help.
'
'The script asks you for a computer name, and will ping the remote
' computer via WMI, then copy three modified help center htm files to the
' remote station (as well as the local station) which allows the unsolicited
' connection to be made.  On the remote user's side, the updated htm file
' will automatically select the 'allow control' option for the remotely
' connected expert.
'
'Once the connection is made, the expert can click 'Take Control' without
' prompting the user to do anything further.
'
'Once the expert has finished with the session to the remote machine,
' an original copy of the htm files are sent back to the computer.
'
'*** Credits:
'
'http://dgrundel.wordpress.com/2007/10/04/unsolicited-remote-assistance/
'http://www.lewisroberts.com/?p=40
'
'This script was primarily written by 'DGrundel', and the changes to the
' HTM files were made by 'DGrundel' and Lewis Roberts.  I made some
' changes to the vbscript below that allows for systems that have Windows
' installed in a folder other than the c:\ drive.
'
'Modifed by: Rob Dunn
'Email: uphold at (two-thousand-one) @ hotmail dot com
'Website: http://www.vbshf.com
'
'Some other tweaks by Kevin Bumber (co-worker) - Thanks!
'
'*** What you need to get started:
'
'In addition to this script, you need the following modified files:
' * helpeeaccept.htm
' * TakeControlMsgs.htm
' * UnSolicitedRCUI.htm
'
'You can copy the originals (search your %windir%\pchealth folders for them)
' and modify them per the instructions specified at DGrundel's website:
' http://dgrundel.wordpress.com/2007/10/04/unsolicited-remote-assistance/
'
'You must modify the following variables within the script to reflect where
' your modified htm files are stored:
'
' * strOriginalFiles
' * strCustomFiles
'
'Note that you must not use leading or trailing backslashes in the path
' as it uses the script's path as a starting point (relative path).
'
'*** Limitations and Security requirements:
'
'If no one is logged into the remote computer at the time when a person
' attempts a connection, the remote assistance will abort - - it needs
' a person on the other end (at least logged in currently) to work.
'
'The person requesting access to the remote station must have administrative
' rights on the remote computer and WMI database (enabled by default)
'
'WMI must be enabled on the remote station (enabled by default)
'
'The firewall must allow remote assistance connections.  If you wish to
' implement this script in a company environment for support purposes,
' you must ensure that the 'Offer Remote Assistance' policy setting
' is enabled and applied to the computer(s) you wish to control.
'
'*** Considerations:
'
'If the connection to the remote computer is broken before the 'Expert'
' finishes the session, the original htm Help Center files are never
' re-copied back to the remote computer, thus opening up a potential
' security vulnerability.  If this happens, you should attempt to
' re-connect to the remote computer again with the script, disconnect
' immediately, and the original htm files will be copied over as a result.
'
'This last point is something I am considering on improving, potentially
' with an automated scheduled task.

'***********************************************************************
'                  Start User Variables
'***********************************************************************

'Relative path to original files (no leading or trailing backslashes) from
' where this script is located.
strOriginalFiles = "Data\Original"

'Relative path to custom files (no leading or trailing backslashes) from
' where this script is located.
strCustomFiles = "Data\Custom"

'**********************************************************************
'                 End User Variables
'**********************************************************************


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
Dim strRemoteDrive, LocalWindir, RemoteWindir, sCommand, strHost

Set objArgs = WScript.Arguments

'*******************
'Get command-line arguments
If objargs.count <> 0 Then
 For I = 0 to objArgs.Count - 1
   If InStr(1,LCase(objargs(I)),"computer:") Then
      arrComputer = split(lcase(objargs(I)),"computer:")
     	strHost = arrComputer(1)
   End If
  Next
Else
  strHost = InputBox("Enter host name or IP Address")
End If


'Verify a host name was actually entered.
If Len(strHost) > 0 Then
  'Ping the host
  If PingStatus(strHost) <> "Success" Then
    MsgBox "Error: Unable to ping " & strHost
    WScript.Quit
  End If

  localOS = GetLocalOS()
  remoteOS = GetRemoteOS()

  'Edited by BUMBER!!!!
  If instr(localOS, "Windows 7") > 0 and instr(remoteOS, "Windows 7") > 0 then   
      ConnectFromNonXP(strHost)
  ElseIf instr(localOS,"Windows 7") then
      ConnectFromNonXP(strHost)
  Else
      strRemoteDrive = replace(GetDrive(strHost),":","")
      strLocalDrive = GetDrive(".")
  
      ConnectFromXP(strHost)
  End If
End If

Function ConnectFromNonXP(strHost)
  sCommand = "msra.exe /offerra " & strHost
  Set WshShell = Wscript.CreateObject("WScript.Shell")
  WshShell.Run(sCommand) 
End Function

Function ConnectFromXP(strHost)
  'Verify you can connect to C$
  If Not objFSO.FolderExists("\\" & strHost & "\" & strRemoteDrive & "$") Then
    MsgBox "Error: Unable to access \\" & strHost & "\" & strRemoteDrive & "$"
    WScript.Quit
  End If

  strScriptPath = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName,"\"))
  
  'Copying custom (security removed) helpeeaccept.htm to the remote machine.
  objFSO.CopyFile strScriptPath & strCustomFiles & "\helpeeaccept.htm","\\" & strHost & "\" & strRemoteDrive & "$\Windows\pchealth\helpctr\System\Remote Assistance\helpeeaccept.htm",True
  
  'Copying custom (security removed) TakeControlMsgs.htm to the remote machine.
  objFSO.CopyFile strScriptPath & strCustomFiles & "\TakeControlMsgs.htm","\\" & strHost & "\" & RemoteWindir & "\pchealth\helpctr\System\Remote Assistance\Interaction\Server\TakeControlMsgs.htm",True

  'Opening custom UnSolicitedRCUI.htm and reading the file into a variable so we can customize the file.
  Set objFile = objFSO.OpenTextFile(strScriptPath & strCustomFiles & "\UnSolicitedRCUI.htm", 1)
  strText = objFile.ReadAll
  objFile.Close

  'Modifying the contents of the file so that it includes the desired host name.
  strNewText = Replace(strText, "idComputerName.value = ""CHANGEME"";", "idComputerName.value =""" & strHost & """;")

  'Writing the modified, custom UnSolicitedRCUI.htm to the local machine.
  Set objFile = objFSO.OpenTextFile(LocalWindir & "\PCHealth\HelpCtr\Vendors\CN=Microsoft Corporation,L=Redmond,S=Washington,C=US\Remote Assistance\Escalation\Unsolicited\UnSolicitedRCUI.htm", 2, True)
  objFile.Write strNewText
  objFile.Close

  'Using WMI to run the modified, custom UnSolicitedRCUI.htm and get the process ID for later use.
  strComputer = "."
  Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
  objWMIService.Create LocalWindir & "\PCHealth\HelpCtr\Binaries\helpctr.exe -url " & "hcp://CN=Microsoft%20Corporation,L=Redmond,S=Washington,C=US/Remote%20Assistance/Escalation/unsolicited/UnSolicitedRCUI.htm", null, null, intProcessID

  'Setting up an event notification to alert the script whenever a new process is created,
  'so we know when the process we created creates a child process (which will be the remote
  'assistance interface.)
  strComputer = "."

  Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
  Set colItems = objWMIService.ExecNotificationQuery("Select * From __InstanceCreationEvent Within 1 Where TargetInstance ISA 'Win32_Process'")

  'Script stays in this loop until the process above creates a child process.
  Do
    Set objProcess = colItems.NextEvent 'Script halts here until a new process is created.

    If objProcess.TargetInstance.ParentProcessId = intProcessID Then 'Checks the Parent PID of the created process to see if its the child we want.
      intChildPID = objProcess.TargetInstance.ProcessId 'Store the process ID of the child process so we monitor it below.
      Exit Do
    End If
  Loop

  'Now that our original process has created a child process, we can terminate it.
  'The code below loops through all processes with image name "helpctr.exe" and
  'terminates the one with our original ProcessID.
  Set colItems = objWMIService.ExecQuery("Select * From Win32_Process Where Name = 'helpctr.exe'")
  For Each objProcess In colItems
    If objProcess.ProcessId = intProcessID Then
      objProcess.Terminate
    End If
  Next

  'Now we monitor for the child process to be closed manually by the remote assistance(RA) expert.
  Set colItems = objWMIService.ExecNotificationQuery("Select * From __InstanceDeletionEvent Within 1 Where TargetInstance ISA 'Win32_Process'")

  'Script stays in this loop until the child process is closed.
  Do
    Set objProcess = colItems.NextEvent 'Script halts here until a process is terminated.
    If objProcess.TargetInstance.ProcessID = intChildPID Then 'Checks to see if the terminated process was our child process.
      Exit Do
    End If
  Loop

  'MsgBox "Click OK when you are finished assisting the remote machine."
  On Error Resume Next

  'When the RA expert disconnects from the remote machine, the remote user sees a message that the expert
  'has disconnected and also is left with the RA chat interface window open. The code below kills helpctr.exe
  'on the remote machine, so the user on the remote machine doesn't have to.
  strComputer = strHost
  Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
  Set colItems = objWMIService.ExecQuery("Select * From Win32_Process Where Name = 'helpctr.exe'")
  For Each objProcess In colItems
    objProcess.Terminate
  Next

  'Replacing the custom helpeeaccept.htm we copied in the beginning with the secure original that came with the OS.
  objFSO.CopyFile strScriptPath & strOriginalFiles & "\helpeeaccept.htm","\\" & strHost & "\" & RemoteWindir & "\pchealth\helpctr\System\Remote Assistance\helpeeaccept.htm",True

  'Replacing the custom TakeControlMsgs.htm we copied in the beginning with the secure original that came with the OS.
  objFSO.CopyFile strScriptPath & strOriginalFiles & "\TakeControlMsgs.htm", "\\" & strHost & "\" & RemoteWindir & "\pchealth\helpctr\System\Remote Assistance\Interaction\Server\TakeControlMsgs.htm", True
  
  'Replacing the custom UnSolicitedRCUI.htm on the local machine with the original that came with the OS.
  objFSO.CopyFile strScriptPath & strOriginalFiles & "\UnSolicitedRCUI.htm", LocalWindir & "\PCHealth\HelpCtr\Vendors\CN=Microsoft Corporation,L=Redmond,S=Washington,C=US\Remote Assistance\Escalation\Unsolicited\UnSolicitedRCUI.htm",True


End Function 'ConnectFromXP


Function PingStatus(strComputer) 'PingStatus function from http://www.microsoft.com/technet/scriptcenter/resources/scriptshop/shop1205.mspx
  On Error Resume Next
  strWorkstation = "."
  Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strWorkstation & "\root\cimv2")
  Set colPings = objWMIService.ExecQuery("SELECT * FROM Win32_PingStatus WHERE Address = '" & strComputer & "'")
  For Each objPing in colPings
    Select Case objPing.StatusCode
      Case 0 PingStatus = "Success"
      Case 11001 PingStatus = "Status code 11001 - Buffer Too Small"
      Case 11002 PingStatus = "Status code 11002 - Destination Net Unreachable"
      Case 11003 PingStatus = "Status code 11003 - Destination Host Unreachable"
      Case 11004 PingStatus = "Status code 11004 - Destination Protocol Unreachable"
      Case 11005 PingStatus = "Status code 11005 - Destination Port Unreachable"
      Case 11006 PingStatus = "Status code 11006 - No Resources"
      Case 11007 PingStatus = "Status code 11007 - Bad Option"
      Case 11008 PingStatus = "Status code 11008 - Hardware Error"
      Case 11009 PingStatus = "Status code 11009 - Packet Too Big"
      Case 11010 PingStatus = "Status code 11010 - Request Timed Out"
      Case 11011 PingStatus = "Status code 11011 - Bad Request"
      Case 11012 PingStatus = "Status code 11012 - Bad Route"
      Case 11013 PingStatus = "Status code 11013 - TimeToLive Expired Transit"
      Case 11014 PingStatus = "Status code 11014 - TimeToLive Expired Reassembly"
      Case 11015 PingStatus = "Status code 11015 - Parameter Problem"
      Case 11016 PingStatus = "Status code 11016 - Source Quench"
      Case 11017 PingStatus = "Status code 11017 - Option Too Big"
      Case 11018 PingStatus = "Status code 11018 - Bad Destination"
      Case 11032 PingStatus = "Status code 11032 - Negotiating IPSEC"
      Case 11050 PingStatus = "Status code 11050 - General Failure"
      Case Else PingStatus = "Status code " & objPing.StatusCode & " - Unable to determine cause of failure."
    End Select
  Next
  On Error Goto 0
End Function

'Get the version of OS for the local computer
Function GetLocalOS()
   Set oWMIService = GetObject("winmgmts:\\.\root\CIMV2")
   Set colItems = oWMIService.ExecQuery("SELECT Caption FROM Win32_OperatingSystem")
   For Each oItem In colItems
        GetLocalOS = oItem.caption
   Next
End Function

Function GetRemoteOS()
   Set oWMIService = GetObject("winmgmts:\\" & strHost &"\root\CIMV2")
   Set colItems = oWMIService.ExecQuery("SELECT Caption FROM Win32_OperatingSystem")
   For Each oItem In colItems
        GetRemoteOS = oItem.caption
   Next
End Function

'Get the local windows folder.
Function GetDrive(sComputer)
   Set oWMIService = GetObject("winmgmts:\\" & sComputer & "\root\CIMV2")
   Set colItems = oWMIService.ExecQuery("SELECT Caption, SystemDrive, WindowsDirectory FROM Win32_OperatingSystem")
   For Each oItem In colItems
        GetDrive = oItem.SystemDrive
        If sComputer = "." Then
          LocalWindir = oItem.WindowsDirectory
        Else
          RemoteWindir = replace(oItem.WindowsDirectory,":","$")
        End if
   Next
End Function

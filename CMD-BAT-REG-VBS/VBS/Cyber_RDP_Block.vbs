' START VBSCRIPT
' **********************************************************
' * Cyber_RDP_Block.vbs
' *
' * Stop Brute Force Dictionary RDP attacks on the fly by null routing attacking IP addresses.
' * This script function on Windows 2003 Server.
' *
' * Note:  Must run continuously:  On Server 2003, create a scheduled task that starts every hour
' *   and terminates itself every 58 minutes (get script rerunning asap on reboots).
' *
' * Note:  Requires event log monitoring of logon events (SECURITY EVENT ID 529).
' *   (Policy\Computer Configuration\Windows Settings\Security Settings\Local Policies\Audit Policy\)
' *   Audit account logon events = (Success,Failure) or (Failure)
' *
' * [CMD] C:\>ROUTE PRINT 'will give you current active routes; show which ones are null routed
'
' * Function:
' *   DO
' *     Monitor the event log for MY_RECORD_SIZE of failed logon attempts
' *     and record event information in Arrays
' *     If MY_THRESHOLD of failed attempts is noted in MY_SECONDS or less then THRESHOLD EXCEEDED:
' *       If last failed SourceIP has MY_THRESHOLD entries in MyArray then ATTACK DETECTED:
' *         If SourceIP detected inside the trusted subnet then ACTIONS:
' *           Report only, do not block
' *           start cooldown on trusted SourceIP reporting
' *           initialize counters to start fresh event detection
' *         Else ACTIONS:
' *           Block SourceIP and report
' *           ' non-persistent - purged on server reboot, helps regulate table size, recommended
' *           ' persistent - retained on reboot, table grows, may require maintenance at some point
' *           initialize counters to start fresh event detection
' *         End Actions
' *       End Attack Detected
' *     End Threshold Exceeded
' *   Loop
' * End Function
' *
' * Written:  James Anderson  Aug 2012
' **********************************************************
' Settings
' MY_RECORD_SIZE must be greater than MY_THRESHOLD.  Record size is how many events are cached.
' Recommended MY_RECORD_SIZE >= 3x MY_THRESHOLD to handle failed logons from multiple IP's
' dimension arrays to MY_RECORD_SIZE + 1
Const MY_RECORD_SIZE = 150
Dim arrMyEvents1(151), arrMyEvents2(151), arrMyEvents3(151), arrMyEvents4(151)
'
' Attack detected when MY_THRESHOLD events detected in MY_SECONDS or less.
Const MY_THRESHOLD = 15
Const MY_SECONDS = 120
'
' Set Trusted Internal network prefix (any IP that starts with this string will be reported but not blocked)
Const MY_TRUSTED_SUBNET = "10.10."
' Set Cooldown in seconds on reporting trusted IP attacks (prevent being flooded with emails)
Const MY_COOLDOWN = 600
'
' Set Null Gateway (unused IP on the subnet) and if you want persistent route entries (permanent ip block)
Const MY_NULL_GATEWAY = "10.10.1.254"
Const MY_ROUTE_PERSISTENT = TRUE
'Const MY_ROUTE_PERSISTENT = FALSE
'
' Reporting Settings 
Const MY_REPORTBYEMAIL = TRUE
'Const MY_REPORTBYEMAIL = FALSE 
Const MY_EMAIL_SERVER = "exchange.mydomain.com"
Const MY_EMAIL_SENDER = "Alert@mydomain.com"
Const MY_EMAIL_RECIPIENT = "MyHelpdesk@mydomain.com;MyManagers@mydomain.com"
Const MY_EMAIL_SUBJECT = "Alert!  RDP cyber attack from IP:  "
'
' **********************************************************
' TEST MODE SETTINGS
Const TEST_MODE = FALSE
'Const TEST_MODE = TRUE
' Sample of IP's that will be presented one by one as the Source Network Address in test mode (; delimited)
strTestIPS = "" _
  & "10.69.254.1;10.69.254.1;10.69.254.2;10.69.254.2;10.69.254.3;10.69.254.3;10.69.254.4;10.69.254.4;" _
  & "10.69.254.1;10.69.254.1;10.69.254.2;10.69.254.2;10.69.254.3;10.69.254.3;10.69.254.4;10.69.254.4;" _
  & "10.69.254.1;10.69.254.1;10.69.254.2;10.69.254.2;10.69.254.3;10.69.254.3;10.69.254.4;10.69.254.4;" _
  & "10.69.254.1;10.69.254.1;10.69.254.2;10.69.254.2;10.69.254.3;10.69.254.3;10.69.254.4;10.69.254.4;" _
  & "10.69.254.1;10.69.254.1;10.69.254.1;10.69.254.1;10.69.254.1;10.69.254.1;10.69.254.1;10.69.254.1;" _
  & "10.69.254.1"
' Sample Test Message generated from a Windows 3003 Server
TEST_MESSAGE = "Microsoft (R) Windows Script Host Version 5.6" & vbcrlf _
  & "Copyright (C) Microsoft Corporation 1996-2001. All rights reserved." & vbcrlf & vbcrlf _
  & "Logon Failure:" & vbcrlf & vbcrlf _
  & vttab & "Reason:" & vbTab & vbTab & "Unknown user name or bad password" & vbcrlf & vbcrlf _
  & vbtab & "User Name:" & vbTab & "MyBadPerson" & vbcrlf & vbcrlf _
  & vbtab & "Domain:" & vbTab & vbTab & "MyDomain" & vbcrlf & vbcrlf _
  & vbtab & "Logon Type:" & vbTab & "7" & vbcrlf & vbcrlf _
  & vbtab & "Logon Process:" & vbTab & "User32  " & vbcrlf & vbcrlf _
  & vbtab & "Authentication Package:" & vbTab & "Negotiate" & vbcrlf & vbcrlf _
  & vbtab & "Workstation Name:" & vbTab & "MyRDPServer" & vbcrlf & vbcrlf _
  & vbtab & "Caller User Name:" & vbTab & "MyRDPServer$" & vbcrlf & vbcrlf _
  & vbtab & "Caller Domain:" & vbTab & "MyDomain" & vbcrlf & vbcrlf _
  & vbtab & "Caller Logon ID:" & vbTab & "(0x0,0x3E7)" & vbcrlf & vbcrlf _
  & vbtab & "Caller Process ID:" & vbTab & "22976" & vbcrlf & vbcrlf _
  & vbtab & "Transited Services:" & vbTab & "-" & vbcrlf & vbcrlf _
  & vbtab & "Source Network Address:" & vbTab & "10.10.10.10" & vbcrlf & vbcrlf _
  & vbtab & "Source Port:" & vbTab & "2567" & vbcrlf & vbcrlf
'
' **********************************************************
' **********************************************************
' MAIN
' initialize test IP's
dim arrTestIPS
arrTestIPS = Split(strTestIPS,";")
inttest = 0
' initialize
strComputer = "."
Const GET_USERNAME = "User Name:"
Const GET_DOMAIN = "Domain:"
Const GET_SOURCEIP = "Source Network Address:"
Const ForReading = 1, ForWriting = 2, ForAppending = 8
i = 1
intTestRecord = 1
intCount = 1
dtCooldown = now()
' initialize the event log monitor to collect new events as they occur
If TEST_MODE = FALSE Then
  Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate, (Security)}!\\" & strComputer & "\root\cimv2")
  Set colMonitoredEvents = objWMIService.ExecNotificationQuery _    
    ("Select * from __instancecreationevent where " _
    & "TargetInstance isa 'Win32_NTLogEvent' " _
    & "and TargetInstance.EventCode = '529' ")
End If
' **********************************************************
' DO LOOP with no exit
Do
  ' Get Event
  If TEST_MODE <> TRUE then
    ' Wait for the next failed logon attempt
    Set objEvent = colMonitoredEvents.NextEvent
  End If
  If intCount < MY_THRESHOLD + 1 Then
    ' have not recorded MY_THRESHOLD events since counter start/reset, record event only
    SrRecordEvent i
  Else
    SrRecordEvent i
    j = i - MY_THRESHOLD
    If j < 1 Then
      j = j + intCount
    End If
    ' Check timestamp difference across MY_THRESHOLD events
    If datediff("s",arrMyEvents1(j),arrMyEvents1(i)) < MY_SECONDS Then
      ' Count the number of times the Source Network Address arrMyEvents2(i) shows up in arrMyEvents
      intDetect = 0
      For j = 1 to intCount
        If  arrMyEvents2(j) = arrMyEvents2(i) Then
          intDetect = intDetect + 1
        End If
      Next
      ' Check arrMyEvents for MY_THRESHOLD of events
      If (intDetect > MY_THRESHOLD) Then
        If TEST_MODE = TRUE Then
          wscript.echo "ATTACK DETECTED ATTACK DETECTED ATTACK DETECTED ATTACK DETECTED ATTACK DETECTED"
        End If
        ' Take Action based on if Source Network Address is on the Trusted Subnet or not
        If Left(arrMyEvents2(i),Len(MY_TRUSTED_SUBNET)) = MY_TRUSTED_SUBNET Then
          'Trusted IP, Report only
          strReport = "RDP logon attack Detected from IP:  " & arrMyEvents2(i) & vbcrlf _
            & "Time:  " & arrMyEvents1(i) & vbcrlf _
            & MY_THRESHOLD & " failed RDP logon attempts detected within " & MY_SECONDS & " seconds." & vbcrlf _
            & "This is an internal subnet IP, no blocking action taken." & vbcrlf & vbcrlf
          If datediff("s",dtCooldown,now()) < MY_COOLDOWN Then
            ' In Cooldown for a spamming Trusted IP
            strReport = "STILL IN COOLDOWN!!!  Remaining seconds:  " & (MY_COOLDOWN - datediff("s",dtCoolDown,now()))
            If TEST_MODE = TRUE Then
              wscript.echo "TRUSTED IP DETECTED:" & arrMyEvents2(i) & vbcrlf _
                & "Report:" & vbcrlf & strReport
            End If
          Else
            If TEST_MODE = TRUE Then
              wscript.echo "TRUSTED IP DETECTED:  " & arrMyEvents2(i) & vbcrlf _
                & "Report:" & vbcrlf & strReport
            Else    
              ' Reporting
              If MY_REPORTBYEMAIL = TRUE Then
                Set objEmail = CreateObject("CDO.Message")
                objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
                objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MY_EMAIL_SERVER
                objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
                objEmail.Configuration.Fields.Update
                objEmail.From = MY_EMAIL_SENDER
                objEmail.To = MY_EMAIL_RECIPIENT
                objEmail.Subject = MY_EMAIL_SUBJECT & arrMyEvents2(i)
                objEmail.TextBody = strReport
                objEmail.send
                wscript.sleep 2000
                set objEmail = nothing
              End If
            End If
            ' start new cooldown for trusted IP, reset counters for fresh search
            dtCoolDown = now()
            intCount = 0
            i = 0
          End If
        Else
          'Not Trusted IP, Take Blocking action.  IP = arrMyEvents2(i)
          If MY_ROUTE_PERSISTENT = FALSE then
            strCommand = "ROUTE ADD " & arrMyEvents2(i) & " MASK 255.255.255.255 " & MY_NULL_GATEWAY & " METRIC 1"
          Else
            strCommand = "ROUTE -p ADD " & arrMyEvents2(i) & " MASK 255.255.255.255 " & MY_NULL_GATEWAY & " METRIC 1"
          End If
          strReport = "Alert!" & vbcrlf _
            & "RDP logon attack Detected from IP:  " & arrMyEvents2(i) & vbcrlf _
            & "Time:  " & arrMyEvents1(i) & vbcrlf _
            & MY_THRESHOLD & " failed RDP logon attempts detected within " & MY_SECONDS & " seconds." & vbcrlf _
            & "Preventative measures have automatically been taken." & vbcrlf _
            & "Command Executed:  " & strCommand & vbcrlf _
            & vbcrlf
          If TEST_MODE = TRUE Then
            wscript.echo "TEST MODE Attack Action:" & vbcrlf _
              & "Command:  " & strCommand & vbcrlf _
              & "Report:" & vbcrlf & strReport
          Else
            ' Execute Command to Null Route the IP address
            Set objCMD = wscript.CreateObject("wscript.shell")
            strReturnCode = objCMD.Run(strCommand,0,true)
            Set objCmd = nothing
            ' Reporting
            If MY_REPORTBYEMAIL = TRUE Then
              Set objEmail = CreateObject("CDO.Message")
              objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
              objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MY_EMAIL_SERVER
              objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
              objEmail.Configuration.Fields.Update
              strReport = strReport & vbcrlf & vbcrlf _
                & "-------Time-------" & vbtab _
                & "--Source IP--" & vbtab _
                & "--User Name--" & vbcrlf
              For j = 1 to intCount
                k = i - j + 1
                If k < 1 Then
                  k = k + intCount
                End If
                If  arrMyEvents2(k) = arrMyEvents2(i) Then
                  strReport = strReport _
                    & arrMyEvents1(k) & vbtab _
                    & arrMyEvents2(k) & vbtab _
                    & arrMyEvents3(k) & vbcrlf
                End If
              Next
              objEmail.From = MY_EMAIL_SENDER
              objEmail.To = MY_EMAIL_RECIPIENT
              objEmail.Subject = MY_EMAIL_SUBJECT & arrMyEvents2(i)
              objEmail.TextBody = strReport
              objEmail.Send
              wscript.sleep 2000
              Set objEmail = nothing
            End if
          End If
          ' action taken, reset counters for fresh search
          intCount = 0
          i = 0
        End If
      End If
    End If
  End If
  ' Event processed, increment counters
  i = i + 1
  If i > MY_RECORD_SIZE Then
    i = i - MY_RECORD_SIZE
  End If
  If intCount < MY_RECORD_SIZE Then
    intCount = intCount + 1
  End If
Loop
' END MAIN
'
Sub SrRecordEvent(i)
  arrMyEvents1(i) = now()
  If TEST_MODE = TRUE Then
    arrMyEvents3(i) = FnParse(TEST_MESSAGE, GET_USERNAME)
    arrMyEvents4(i) = FnParse(TEST_MESSAGE, GET_DOMAIN)
    arrMyEvents2(i) = arrTestIPS(inttest)
    inttest = inttest + 1
    If inttest > ubound(arrTestIPS) then
      inttest = 0
    End If
    wscript.echo "TEST MODE SIMULATE FAILED LOGON!" & vbcrlf _
      & "Buffer:  " & intCount & "  RecordEvent(" & i & ") as " _
      & vbcrlf & arrMyEvents1(i) & "|" & arrMyEvents3(i) & "|" & arrMyEvents4(i) & "|" & arrMyEvents2(i)
  Else
    arrMyEvents2(i) = FnParse(objEvent.TargetInstance.Message, GET_SOURCEIP)
    arrMyEvents3(i) = FnParse(objEvent.TargetInstance.Message, GET_USERNAME)
    arrMyEvents4(i) = FnParse(objEvent.TargetInstance.Message, GET_DOMAIN)
  End If
End Sub
'
Private Function FnParse(byVal strMessage, strValue)
  ' GET_VALUE in strMessage
  dim tmp1
  Set objRegEx = CreateObject("VBScript.RegExp")
  objRegEx.Global = TRUE
  objRegEx.IgnoreCase = TRUE
  objRegEx.Pattern = "(\t" & strValue & ".*\n)"
  Set colMatch = objRegEx.Execute(strMessage)
  For Each strMatch in colMatch
    strtmp1 = strMatch.Value
  Next
  tmp1 = Split(strtmp1,":")
  FnParse = Trim(Replace(Replace(tmp1(1),vbTab,""),vbCrLf,""))
End Function
' END VBSCRIPT

'~~[author]~~
'Laurent Belloeil
'~~[/author]~~

'~~[emailAddress]~~
'lbelloeil@yahoo.fr
'~~[/emailAddress]~~

'~~[scriptType]~~
'vbscript
'~~[/scriptType]~~

'~~[subType]~~
'SystemAdministration
'~~[/subType]~~

'~~[keywords]~~
'system, Service, Sart, Run, Stop, Error, Backup, Delay, Log, Ie, html
'~~[/keywords]~~

'~~[comment]~~
'A powerfull script (v1.0) to start top service on any computer. Usefull to prepare for a backup. Can start/stop a service on any computer. Verify if service exist, control service before and after action, log in IE (can easily be modified to write to html or text file), return error codes. You have to specify a delay to perform the operation.

'Please send me your comments.
'~~[/comment]~~

'~~[script]~~
'****************************************************************************************************
'*	Script Name:Service manager
'*	Version: 1.0
'*	Purpose:	To allow the stopping and starting of services, with timing and checking.
'*	Inputs:	domain name, computer name, service name, action (sart|stop), delay for action (s)
'*	Example:	SvcMangr.vbs workgroup home wuauserv stop 10
'*	Returns:	The error code where 0 is successful
'*	Future:	Verify if the service exist and its state before action
'*			Display log In IE but you can redirect To a text file
'*	Error codes :	10	syntax error
'*				-2	Computer not found
'*				-1	service not found
'*				1	already stopped
'*				2	starting
'*				3	stopping
'*				4	running
'*				5	still running
'*				6	pausing
'*				7	paused
'*				8	on error
'*				9	undefined state or out of delay
'*	
'****************************************************************************************************
Option Explicit

Const STOPPED = 1
Const START_PENDING = 2
Const STOP_PENDING = 3
Const RUNNING = 4
Const CONTINUE_RUNNING = 5
Const PAUSE_PENDING = 6
Const PAUSED = 7
Const Error = 8

'Objects and variables definitions
Dim objComputer, objNTService, objIE, WSHNetwork, objWshScriptExec, objShell ,objStdOut
Dim strText, strStatus, strDomain, strComputer, strService, strAction, hostName, strLine, strIETitle
Dim i, j, valDelay, nbsec, RetCode, intWidth, intWidthW, intHeight, intHeightW
On Error Resume Next

'Verify and assign arguments to named variables
If WScript.arguments.count <> 5 Then
	WScript.Echo "Usage: SvcMangr.Vbs <domainname> <computername> <servicename> start|stop <delai (s)>" 
	WScript.quit (10)
End If
strDomain = lcase (WScript.arguments(0))
strComputer = lcase (WScript.arguments(1))
strService = lcase (WScript.arguments(2))
strAction = lcase (WScript.arguments(3))
valDelay = cint(WScript.arguments(4))
strIETitle = "Service Manager"
Set objShell = CreateObject("WScript.Shell")

	'- make changes here in order to write logs to a file ----------
	'Set the properities For the IE Object
	Set objIE = WScript.CreateObject("InternetExplorer.Application")
	objIE.Navigate "about:blank"
	objIE.Visible = 1
	objIE.Document.Body.InnerHTML = strText

	  With objIE
	    .ToolBar = False
	    .StatusBar = False
	    .Resizable = False
	    .Navigate("about:blank")
	    Do Until .readyState = 4
	      Wscript.Sleep 100
	    Loop
	    With .document
		      .Title = strIETitle
		      With .ParentWindow
		        intWidth = .Screen.AvailWidth
		        intHeight = .Screen.AvailHeight
		        intWidthW = .Screen.AvailWidth * .40
		        intHeightW = .Screen.AvailHeight * .40
		        .resizeto intWidthW, intHeightW
		        .moveto (intWidth - intWidthW)/2, (intHeight - intHeightW)/2
		      End With
'	      .Write "<body> " & strMsg & " </body></html>"
'	      With .ParentWindow.document.body
'	        .style.backgroundcolor = "LightBlue"
'	        .scroll="no"
'	        .style.Font = "10pt 'Courier New'"
'	        .style.borderStyle = "outset"
'	        .style.borderWidth = "4px"
'	      End With
	      objIE.Visible = True
	      Wscript.Sleep 100
	      objShell.AppActivate (strIETitle)
	    End With
	  End With
	strText = "<font size=" & chr(34) & "2" & chr(34) & "face=" & chr(34) & "Tahoma" & chr(34) & "><P>"
	'---------------------------------------------------------------
	
'Verifiy if computer Is reachable
If blnPingable(strComputer) Then
	strText=strText & "Computer alive..."
	HTMLnextPar(strText)
	'Connect to computer
	Set objComputer=GetObject("WinNT://"+strDomain +"/"+strComputer +",computer") 
			
	'Verify if the service exist
	retCode=-1
	objComputer.Filter=Array("service")
	For Each objNTService In objComputer 
		objIE.Document.Body.InnerHTML = strText & objNTService.Name
		If lcase(objNTService.Name) =strService Then
		
			'Service Found
			strText = strText & "<i>" & strService & "</i> found On computer <i>" & strComputer & "</i>"
			Call HTMLnextPar (strText )
			Select Case objNTService.status
			'Retourne un entier long qui correspond à l'état du service :
				Case STOPPED 
					If strAction = "start" Then
						strText = strText + "Service stopped. Trying to start (" &valDelay&"s): "
						Call HTMLcont (strText )
						objNTService.Start
						For i= 1 To valDelay
							For j=1 To 10
								WScript.Sleep (100 )
							Next
							If objNTService.Status = RUNNING Then
								nbsec = i
								Exit For
							End If
							strText = strText + cstr(i)+","
							Call HTMLcont (strText )
						Next
							
						'Vérify if service is started
						If objNTService.Status = RUNNING Then
							Call HTMLnextPar (strText )
							strStatus = " has been started within "& nbsec &" seconds"
							Retcode = 0
						Else
							strStatus = " has NOT been started within "+cstr(valDelay) +" seconde"
							RetCode = 9
						End If
					Else
						strStatus = " is already stopped"
						retCode = 1
					End If				
				Case START_PENDING 
					strStatus = " is starting"
					RetCode = 2
				Case STOP_PENDING 
					strStatus = " is stopping"
					RetCode = 3
				Case RUNNING
					If strAction = "stop" Then
						strText = strText + "<P> Service running. Trying to stop (" & valDelay & "s): "
						Call HTMLcont (strText )
						objNTService.Stop
						For i= 1 To valDelay
							For j=1 To 10
								WScript.Sleep (100 )
							Next
							If objNTService.Status = STOPPED Then
								nbsec = i
								Exit For
							End If
							strText = strText + cstr(i)
							Call HTMLcont (strText )
						Next
							
						'Vérify if service stopped
						If objNTService.Status = STOPPED Then
							Call HTMLnextPar (strText )
							strStatus = " has been stopped within "& i  &" seconds"
							Retcode= 0
						Else
							strStatus = " has NOT been stopped within "+cstr(valDelay) +" seconds"
							RetCode = 9
						End If
					Else
						strStatus = " IS ALREADY RUNNING"
						RetCode = 4
					End If
				Case CONTINUE_RUNNING 
					strStatus = " STILL RUNNING"
					RetCode = 5
				Case PAUSE_PENDING 
					strStatus = " IS PAUSING"
					RetCode = 6
				Case PAUSED 
					strStatus = " IS PAUSED"
					RetCode = 7
				Case Error 
					strStatus = " IS ON ERROR"
					RetCode = 8
				Case Else 
					strStatus = " IS UNDEFINED STATE"
					RetCode = 9
			End Select
			Exit For
		End If
	Next
Else
    retcode = -2
End If

Select Case RetCode
	Case -2
		strText = strText & "<b>The computer <i>" & strComputer& "</i> is currently UNREACHABLE</b></P>"
		'- make changes here in order to write logs to a file ----------
		Call HTMLnextPar (strText )
		'---------------------------------------------------------------
	Case -1
		strText = strText & "<b>THE SERVICE <i>" & strservice & "</i> HAS Not BEEN FOUND On COMPUTER <i>" & strComputer & "</b></P>"
		'- make changes here in order to write logs to a file ----------
		Call HTMLnextPar (strText )
		'---------------------------------------------------------------
	Case Else
			strText = strText & "The service <i>" & strservice & "</i>" & strStatus & " on computer <i>" & strComputer & "</b></P>"
		'- make changes here in order to write logs to a file ----------
		Call HTMLnextPar (strText )
		'---------------------------------------------------------------
End Select

'Call HTMLcont (strText & "</P><P align=" & chr(34) & "right" & chr(34) & "><dl><dt><dd><input type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Close" & chr(34) & " name=" & chr(34) & "Close" & chr(34) & " onClick=" & chr(34) & "window.self.close()" & chr(34) & "></p)")
wscript.quit (retCode)

Sub HTMLnextLine (TxtToDisp)
	TxtToDisp = TxtToDisp & "<br>"
	objIE.Document.Body.InnerHTML = TxtToDisp
End Sub

Sub HTMLnextPar (TxtToDisp)
	TxtToDisp = TxtToDisp & "</P>" & "<P>"
	objIE.Document.Body.InnerHTML = TxtToDisp
End Sub

Sub HTMLcont (TxtToDisp)
	objIE.Document.Body.InnerHTML = TxtToDisp
End Sub
'************************************************************************************
'Function blnPingable(strHost)
'************************************************************************************

Function blnPingable(strHost)

Dim objExec, objRE

blnPingable = False 'assume failure
Set objExec = objShell.Exec("cmd /C ping -n 2 " & strHost)

Set objRE = New RegExp
objRE.IgnoreCase = True
objRE.Pattern = "TTL="

If objRE.Test(objExec.StdOut.ReadAll) Then blnPingable = True 'only one test for success

Set objExec = Nothing
Set objRE = Nothing

End Function
'------------------------------------------------------------------

'~~[/script]~~

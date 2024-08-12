'sVer = ".1 - John Battista 10/15/2012"
'sModified = "10/15/12 - JBB"
'
'

' ******* Global Statements *******
bDebug = false

if bDebug = false then
on error resume next
Else
on error goto 0
End if

Const dictKey = 1
Const dictItem = 2
Const ForAppending = 8
Const ForWriting = 2
Const ForReading = 1

dim max,min, computer_ou, user_ou, sIPAddress1, sLetter, sPath, sPrintQueue, sDefault
Randomize
max=1000
min=1
sRandomNumber = Int((max-min+1)*Rnd(10)+min)
sRandomNumber2 = Int((max-min+1)*Rnd(10)+min)

Dim bMapDrive, bMapPrinter, bRunInclude

Dim sUser, sComputer, warn, sLogonServer, log, sDomain, WshShell, WshSysEnv, objFSO, objArgs, objNetwork, objGroupList, oShell

WshShell = sRandomNumber & "_" & WshShell
WshSysEnv = sRandomNumber & "_" & WshSysEnv
objFSO = sRandomNumber & "_" & objFSO
objArgs = sRandomNumber & "_" & objArgs
objNetwork = sRandomNumber & "_" & objNetwork
oShell = sRandomNumber & "_" & oShell

Set WshShell = CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("PROCESS")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objArgs = WScript.Arguments
Set objNetwork = CreateObject("WScript.Network")
Set objGroupList = CreateObject("Scripting.Dictionary")

if bDebug = false then
	On error resume next
Else
	On error goto 0
End if

'Get the complete path of our logon script (minus the script name)
strScriptPath = replace(wscript.scriptfullname,wscript.scriptname,"")
sRandomNumber = sRandomNumber2 + sRandomNumber

'Set the log file location
logfile = WshSysEnv("TEMP") & "\" & "logon_" & sRandomNumber & ".log"
'Get the logged in username
sUser = WshSysEnv("username")
'Get the userprofile location
sUserProfile = WshSysEnv("userprofile")
'Get the current computer name
sComputer = WshSysEnv("computername")
'Get the computer's logon server
sLogonServer = WshSysEnv("logonserver")
'Get the computer's logon domain
sDomain = WshSysEnv("userdomain")
'Configurations folder, can accept a relative path from the location where
' this script is run from. Leading "\" is not required.
sConfigFolder = strScriptPath & "configs\"
'Includes folder, can accept a relative path from the location where this
' script is run from. Leading "\" is not required.
sIncludeFolder = strScriptPath & "includes\"
'Set to '1' if you want to give the user notification of errors.
Warn = 0

'open our logfile for writing.
on error resume next

Set log = objFSO.OpenTextFile (logfile, ForWriting, True)

on error resume next

'Get AD computer object via name
Set objADInfo = CreateObject("ADSystemInfo")
strComputer1 = objADInfo.ComputerName
strUser1 = objADInfo.UserName
Set objComputer = GetObject("LDAP://" & strComputer1)
Set objUser = GetObject("LDAP://" & strUser1)

WriteLog("Running modular logon script " & sVer & ": " & wscript.scriptfullname)
WriteLog("Running via: " & wscript.fullname)

if bDebug = false then
	On error resume next
Else
	On error goto 0
End if

WriteLog("Logon server: " & sLogonServer)

If objComputer.Parent <> "" Then
	computer_OU = replace(objComputer.Parent,"LDAP://","")
	WriteLog("Computer OU: " & computer_OU)
End If

If objComputer.Parent <> "" Then
	user_OU = replace(objUser.Parent,"LDAP://","")
	WriteLog("User OU: " & user_OU)
End If

on error goto 0
Dim smini

aIPAddress = split(GetIPAddress(),"|")

'Strip IPv6 addresses from array.
For a = 0 to ubound(aIPAddress)
	amini = split(aIPaddress(a)," - ")

	smini = amini(0) & delim & smini

	delim = ";"
Next

WriteLog("IP Address(es): " & smini)

Dim aIPAddress

aIPAddress = split(smini,";")

'Get the command line arguments given
If objArgs.Count < 0 Then
	WriteLog("No configuration file given, no logon settings will be applied")
Else
	WriteLog("User '" & sUser & "' is logging into computer " & sComputer & " from domain " & sDomain)
	For I = 0 to objArgs.Count - 1
		If objArgs.Count > 0 Then
		'look for specified configuration file
			If instr(LCase(objargs(i)),"config:") Then
				aConfig = split(objargs(0),":")
				sConfig = trim(aConfig(1))
				'if our specified path does not include backslashes, then we must
				' assume it is in our predefined 'configs' folder.

				If Not instr(sConfig,"\") then
					sConfig = trim(sConfigFolder & sConfig)
				End if

				WriteLog("Using config file: " & sConfig)
				on error resume next

				If ReportFileStatus(sConfig) <> "true" then
					WriteLog(sConfig & " not found. Cannot continue.")
					tArray = split(sConfig,"\")
					log.close
					on error resume next
					objFSO.MoveFile logfile, replace(replace(logfile,".log","_" & tArray(ubound(tArray)) & ".log"),sRandomNumber & "_","ERROR_")
					wscript.quit
				End If

				If bDebug = false then
					On error resume next
				Else
					On error goto 0
				End if

				Dim objUser, strDN

				Const ADS_NAME_INITTYPE_GC = 3
				Const ADS_NAME_INITTUPE_DOMAIN = 1
				Const ADS_NAME_TYPE_NT4 = 3
				Const ADS_NAME_TYPE_1779 = 1

				'Translate the domain and username into a proper LDAP object
				Set objTrans = CreateObject("NameTranslate")

				objTrans.Init ADS_NAME_INITTYPE_GC, ""
				objTrans.Set ADS_NAME_TYPE_NT4, sDomain & "\" & sUser

				strDN = objTrans.Get(ADS_NAME_TYPE_1779)
				strDN = Replace(strDN, "/", "\/")

				' Bind to the user object with the LDAP provider.
				' strDN = Wscript.Arguments(0)
				On Error Resume Next
				Set objUser = GetObject("LDAP://" & strDN)

				' Enumerate group memberships.
				Call EnumGroups(objUser)
				Writelog("...checking group memberships for user object...")
				groupcount = 0
				For Each groupName In objGroupList
					groupcount = groupcount + 1
					WriteLog(groupName)
				Next
				
				Writelog("User is a member of " & groupcount & " domain groups")

				'ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile(sConfig).ReadAll

				Set objTextFile = objFSO.OpenTextFile(sConfig, ForReading)

				strText = lcase(objTextFile.ReadAll)

				strText = replace(replace(strText,"suser",sUser),"scomputer",sComputer)
				strText = replace(replace(strText,"suserprofile",sUserProfile),"slogonserver",sLogonServer)
				strText = replace(replace(replace(strText,"sdomain",sDomain),"user_ou",User_OU),"computer_ou",Computer_OU)

				if bDebug = false then
					On error resume next
				Else
					On error goto 0
				End if

				on error goto 0
				Call ProcessConfig(strText)

				if err.number <> 0 then writelog(err.description & " " & err.number & ", " & err.source)
			End If
		End If
	Next
End If

Public Function ExecuteBasedOnCriteria(sSection,aPerms,bPerms)
on error goto 0
Select Case sSection
Case "include"

if isNull(bPerms) then msgbox "NULL!"

	Call ExecuteScript(sIncludeFolder & aPerms(0))
	
		
	Case "drives"
		Call MapNetworkDrives(sLetter,sPath)
	Case "printers"
		Call AddNetworkPrinter(sPrintQueue,sDefault)
	End Select
End Function 'ExecuteBasedOnCriteria

Public Function RemoveBasedOnCriteria(sSection,aPerms,bPerms)
	Select Case sSection
		Case "drives"
			Call RemoveNetworkDrives(sLetter,"(" & sPath & ")")
		Case "printers"
			Call RemoveNetworkPrinter(sPrintQueue)
	End Select
End Function

Public Function CheckCriteria(aPermDetails,sSection,aPerms)

If sSection = "include" then

	sUserDenyVerbage = "User '" & aPermDetails(1) & "' has been explicitely denied executing this code."
	sGroupDenyVerbage = "User is a member of " & aPermDetails(1) & ", but script explicitely denies executing this code for this group's members."
	sUOUDenyVerbage = "User is a member of OU: " & aPermDetails(1) & ", but script explicitely denies mapping this printer for this OU's members."
	sOUDenyVerbage = "Computer is a member of OU: " & aPermDetails(1) & ", but script explicitely denies executing this code for this OU's members."
	sIPDenyVerbage = "IP string '" & aPermDetails(1) & "' is part of this computer's IP address; This script explicitely denies executing process '" & sIncludeFolder & aPerms(0) & "' for this IP criteria."
	sComputerDenyVerbage = "Computer '" & aPermDetails(1) & " ' has been explicitely denied executing this code."
	sActionVerbage = "Executing " & sIncludeFolder & aPerms(0)
	sNotActionVerbage = "Not executing " & sIncludeFolder & aPerms(0)

ElseIf sSection = "drives" Then
	sOUDenyVerbage = "Computer is a member of OU: " & aPermDetails(1) & ", but script explicitely denies mapping this drive for this OU's members."
	sUserDenyVerbage = "User '" & aPermDetails(1) & "' has been explicitely denied mapping this drive."
	sGroupDenyVerbage = "User is a member of " & aPermDetails(1) & ", but script explicitely denies mapping this drive for this group's members."
	sUOUDenyVerbage = "User is a member of OU: " & aPermDetails(1) & ", but script explicitely denies mapping this drive for this OU's members."
	sIPDenyVerbage = "IP string '" & aPermDetails(1) & "' is part of this computer's IP address; This script explicitely denies mapping this drive '" & sLetter & sPath & "' for this IP criteria."
	sComputerDenyVerbage = "Computer '" & aPermDetails(1) & " ' has been explicitely denied mapping this drive."
	sActionVerbage = "Mapping drive " & sLetter
	sNotActionVerbage = "Not mapping drive " & sLetter

ElseIf sSection = "printers" Then

	sOUDenyVerbage = "Computer is a member of OU: " & aPermDetails(1) & ", but script explicitely denies mapping this printer for this OU's members."
	sUserDenyVerbage = "User '" & aPermDetails(1) & "' has been explicitely denied mapping this printer."
	sGroupDenyVerbage = "User is a member of " & aPermDetails(1) & ", but script explicitely denies mapping this printer for this group's members."
	sUOUDenyVerbage = "User is a member of OU: " & aPermDetails(1) & ", but script explicitely denies mapping this printer for this OU's members."
	sIPDenyVerbage = "IP string '" & aPermDetails(1) & "' is part of this computer's IP address; This script explicitely denies mapping printer '" & sPrintQueue & "' for this IP criteria."
	sComputerDenyVerbage = "Computer '" & aPermDetails(1) & " ' has been explicitely denied mapping this printer."
	sActionVerbage = "Mapping printer " & sPrintQueue
	sNotActionVerbage = "Not mapping printer '" & sPrintQueue

End If

	If lcase(right(aPermDetails(0),10)) = "computerou" then
		If lcase(left(aPermDetails(0),1)) = "-" and inOU(aPermDetails(1),"Computer",false) = true then
			Call WriteLog(sOUDenyVerbage)
			bPerms = False
			Call RemoveBasedOnCriteria(sSection,aPerms,bPerms)
		ElseIf lcase(left(aPermDetails(0),1)) <> "-" and inOU(aPermDetails(1),"Computer",true) = true then
			WriteLog(sActionVerbage)
			bPerms = True
			Call ExecuteBasedOnCriteria(sSection,aPerms,bPerms)
		Else
			bPerms = NULL
			Call WriteLog(sNotActionVerbage)
		End If

	ElseIf lcase(right(aPermDetails(0),2)) = "ip" then
		For ipnum = 0 to ubound(aIPAddress)
			If instr(aIPAddress(ipnum),aPermDetails(1)) then
				bIP = true
			End if
		Next

		If lcase(left(aPermDetails(0),1)) = "-" and bIP = true then
			Call WriteLog(sIPDenyVerbage)
			bPerms = False
			Call RemoveBasedOnCriteria(sSection,aPerms,bPerms)
		ElseIf lcase(left(aPermDetails(0),1)) <> "-" and bIP = true then
			Call WriteLog(sActionVerbage)
			bPerms = true
			Call ExecuteBasedOnCriteria(sSection,aPerms,bPerms)
		Else
			bPerms = NULL
			Call WriteLog(sNotActionVerbage)
		End If

	ElseIf lcase(right(aPermDetails(0),10)) = "userou" then
		If lcase(left(aPermDetails(0),1)) = "-" and inOU(aPermDetails(1),"User",false) = true then
			Call WriteLog(sUOUDenyVerbage)
			Call RemoveBasedOnCriteria(sSection,aPerms,bPerms)
		ElseIf lcase(left(aPermDetails(0),1)) <> "-" and inOU(aPermDetails(1),"User",true) = true then
			WriteLog(sActionVerbage)
			Call ExecuteBasedOnCriteria(sSection,aPerms,bPerms)
		Else
			Call WriteLog(sNotActionVerbage)
		End If

	ElseIf lcase(right(aPermDetails(0),5)) = "group" then
		If lcase(left(aPermDetails(0),1)) = "-" and IsMember(aPermDetails(1),false) = true then
			Call WriteLog(sGroupDenyVerbage)
			bPerms = False

			Call RemoveBasedOnCriteria(sSection,aPerms,bPerms)
		ElseIf lcase(left(aPermDetails(0),1)) <> "-" and IsMember(aPermDetails(1),true) = true then
			WriteLog(sActionVerbage)
			bPerms = True

			Call ExecuteBasedOnCriteria(sSection,aPerms,bPerms)
		Else
			bPerms = NULL

			WriteLog(sNotActionVerbage)
		End If

	ElseIf lcase(right(aPermDetails(0),4)) = "user" then
		If lcase(left(aPermDetails(0),1)) = "-" and lcase(sUser) = lcase(aPermDetails(1)) then
			Call WriteLog(sUserDenyVerbage)
			Call RemoveBasedOnCriteria(sSection,aPerms,bPerms)
		ElseIf lcase(left(aPermDetails(0),1)) <> "-" and lcase(sUser) = lcase(aPermDetails(1)) then
			WriteLog(sActionVerbage)
			Call ExecuteBasedOnCriteria(sSection,aPerms,bPerms)
		End If
		
	ElseIf lcase(right(aPermDetails(0),12)) = "computername" then
		If lcase(left(aPermDetails(0),1)) = "-" and lcase(sComputer) = lcase(aPermDetails(1)) then
			Call WriteLog(sComputerDenyVerbage)
			Call RemoveBasedOnCriteria(sSection,aPerms,bPerms)
		ElseIf lcase(left(aPermDetails(0),1)) <> "-" and lcase(sComputer) = lcase(aPermDetails(1)) then
			WriteLog(sActionVerbage)
			Call ExecuteBasedOnCriteria(sSection,aPerms,bPerms)
		End If
	
	End If
	
End Function 'CheckCriteria

Function ExecuteScript(sScript)

	on error resume next
	ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile(sScript).ReadAll
	if err.number <> 0 then
		call logerror(err.number,err.description,"'ExecuteScript(" & sScript & ")'")
		ExecuteScript = false
	End If

	End Function

Public Function ProcessSection(sSection)
'msgbox sSection

If bDebug = false then
	On error resume next
Else
	On error goto 0
End if

sSelection = replace(sSection,chr(34),"")
WriteLog("======================================================================")
WriteLog(" " & ucase(sSelection))
WriteLog("======================================================================")

'msgbox sSelection
on error goto 0

	Select Case sSelection

		Case "include"

		If instr(lcase(strText),"<include>") then
			sInclude = split(strText,"<include>")
			sTempVar = split(lcase(sInclude(1)),"</include>")
			aIncludeData = split(trim(sTempVar(0)),vbnewline)
		End If

		Dim sInclude

		For i = 1 to ubound(aIncludeData) - 1

			If instr(lcase(aIncludeData(i)),"rem ") and i = (ubound(aIncludeData) - 1) then
				exit Function
			ElseIf instr(lcase(aIncludeData(i)),"rem ") then
				i = i + 1
			End If

			sIncludeData = aIncludeData(i)


			If instr(sIncludeData,"|") then
				aPerms = split(sIncludeData,"|")
				sInclude = aPerms(0)
				aPermsData = aPerms(1)
				'msgbox aPerms(0)
			Else
				aPermsData = ""
				sInclude = sIncludeData
			End If

			on error goto 0
			
			If aPermsData <> "" then

				For j = 1 to ubound(aPerms)

					aPermDetails = split(aPerms(j),":")

					Call CheckCriteria(aPermDetails,"include",aPerms)
					sInclude = ""
					sPermsData = ""

				Next
			Else
				on error resume next
				WriteLog("Executing " & sIncludeFolder & sInclude)
				ExecuteScript(sIncludeFolder & sInclude)
				sInclude = ""
				sPermsData = ""
			End If
		Next

		Case "printers"
			'Process printers
			If instr(lcase(strText),"<printers>") then
				sPrinters = split(strText,"<printers>")
				sTempVar = split(lcase(sPrinters(1)),"</printers>")
				sPrintersData = split(sTempVar(0),vbnewline)
				'msgbox ubound(sPrintersData)

				For i = 1 to ubound(sPrintersData) - 1

					If instr(lcase(sPrintersData(i)),"rem ") and i = (ubound(sPrintersData) - 1) then
						exit function
					ElseIf instr(lcase(sPrintersData(i)),"rem ") then
						i = i + 1
					End If

					sDefault = false
					dim aPerms

					If instr(lcase(sPrintersData(i)),";default") then
						sDefault = true
						sPrintersData1 = replace(lcase(sPrintersData(i)),";default","")
					Else
						sPrintersData1 = sPrintersData(i)
					End If

					If instr(sPrintersData1,"|") then
						aPerms = split(sPrintersData1,"|")
						sPrintQueue = aPerms(0)
					Else
						aPerms = ""
						sPrintQueue = sPrintersData1
					End If
					
					on error resume next

					If not isnull(aperms) and aPerms <> "" then
						For j = 1 to ubound(aPerms)
							aPermDetails = split(aPerms(j),":")
							Call CheckCriteria(aPermDetails,"printers",aPerms)
						Next
					Else
						Call AddNetworkPrinter(sPrintQueue,sDefault)
					End If
				Next
			Else
				writelog("No Printers section found. Skipping.")
			End if

		Case "drives"

			'Process drives
			on error goto 0
			If instr(lcase(strText),"<drives>") then
				sDrives = split(strText,"<drives>")

				sTempVar = split(lcase(sDrives(1)),"</drives>")
				aDrivesData = split(trim(sTempVar(0)),vbnewline)

				'aDrivesData = split(sTempVar(0),vbcrlf)
				bmapped = false

				For i = 1 to ubound(aDrivesData) - 1
					If instr(lcase(aDrivesData(i)),"rem ") and i = (ubound(aDrivesData) - 1) then
						exit function
					ElseIf instr(lcase(aDrivesData(i)),"rem ") then
						i = i + 1
					End If

					on error resume next
					sDrivesData = aDrivesData(i)
					
					If instr(sDrivesData,"|") then
						aPerms = split(sDrivesData,"|")
						sPath = aPerms(0)
						'msgbox aPerms(0)
						aPathDetails = split(sPath,"\\")
					Else
						aPerms = ""
						aPathDetails = split(sDrivesData,"\\")
					End If

					'msgbox sDrivesData
					'msgbox aPathDetails(0)

					sLetter = aPathDetails(0)
					sPath = "\\" & aPathDetails(1)

					If not isnull(aperms) and aPerms <> "" then
						For j = 1 to ubound(aPerms)
							aPermDetails = split(aPerms(j),":")
							Call CheckCriteria(aPermDetails,"drives",aPerms)
						Next
					Else
						Call MapNetworkDrives(sLetter,sPath)
					End If
					bmapped = false
				Next
			Else
				writelog("No Drives section found. Skipping.")
			End if

		Case "processes"

			'Process processes
			If instr(lcase(strText),"<processes>") then
				sProcesses = split(strText,"<processes>")
				sTempVar = split(lcase(sProcesses(1)),"</processes>")
				sProcessesData = split(trim(sTempVar(0)),vbnewline)

				For i = 1 to ubound(sProcessesData) - 1
					'msgbox "running " & sProcessesData(i)
					If sProcessesData(i) <> "" then Call RunProcess(sProcessesData(i),false)
				Next
			Else
				writelog("No Processes section found. Skipping.")
			End if

		Case "meta"

			'Process meta
			If instr(lcase(strText),"<meta>") then
				sMeta = split(lcase(strText),"<meta>")
				sTempVar = split(sMeta(1),"</meta>")
				aMetaData = split(trim(sTempVar(0)),vbnewline)

				For i = 0 to ubound(aMetaData)

					If aMetaData(i) <> "" then
						aTemp = split(aMetaData(i),"=")
						'msgbox aTemp(0) & " value is " & aTemp(1)
						writelog(ucase(aTemp(0)) & " = " & ucase(aTemp(1)))
					End if
				Next
			Else
				writelog("No Meta section found. Skipping.")
				'exit subE
			End if

		Case Else
			'msgbox "TEST!"
		End Select

End Function

Sub ProcessConfig(strText)

	Dim scriptSections ' Create a variable.
	Set scriptSections = CreateObject("Scripting.Dictionary")

	'Process order
	If instr(lcase(strText),"<order>") then
		sOrder = split(strText,"<order>")
		sTempVar = split(lcase(sOrder(1)),"</order>")

		aOrder = split(trim(sTempVar(0)),vbnewline)

		For i = 0 to ubound(aOrder)

			If aOrder(i) <> "" then

				aTemp = split(aOrder(i),"=")
				scriptSections.add chr(34) & aTemp(1) & chr(34), chr(34) & aTemp(0) & chr(34)
				'msgbox aTemp(0) & " value is " & aTemp(1)
				'writelog(aTemp(0) & " = " & aTemp(1))
			End if
		Next
	Else
		writelog("No Order section found.")
	End if

	SortDictionary scriptSections,dictkey

	For Each i In scriptSections
		on error goto 0
		Call ProcessSection(replace(scriptSections(i),chr(34),""))
	Next

End Sub


WriteLog("Logon process completed")

'**** Set declared objects to nothing to free up memory.
Set userObj = nothing
Set objNetwork = nothing
Set oShell = nothing
Set WshSysEnv = nothing
on error resume next

'close our log file.
log.close

tArray = split(sConfig,"\")

'msgbox replace(replace(logfile,".log","_" & tArray(ubound(tArray)) & ".log"), sRandomNumber & "_","")

on error resume next
Set MyFile = objFSO.GetFile(replace(replace(logfile,".log","_" & tArray(ubound(tArray)) & ".log"), sRandomNumber & "_",""))
MyFile.delete

objFSO.MoveFile logfile, replace(replace(logfile,".log","_" & tArray(ubound(tArray)) & ".log"),sRandomNumber & "_","")

'Set MyFile = objFSO.GetFile(logfile)
'MyFile.delete

' ******* End of Global Statements *******
'##############################################################################'

'=============================================================================='
' Functions and Subroutines... '
' You probably won't need to modify these. '

'------------------------------------------------------------------------------'
' Function fctCopyFile -
' Copy a file from strSource (path) to strDestFolder (path)
'##############################################################################'

Function fctCopyFile(strSource,strDestFolder)
	Set MyFile = objfso.GetFile(strSource)
	MyFile.Copy (strDestFolder)
End Function

'------------------------------------------------------------------------------'
' Subroutine Set Attribute - sets 'comment' attribute in logged-on
' user object to computer name.
'##############################################################################'

Sub SetAttribute()
	'Script to put computername into 'comment' attribute on user object in
	' Active Directory.
	'on error resume next

	' Get the NETBIOS Domain name
	Set objSystemInfo = CreateObject("ADSystemInfo")
	sDomain = objSystemInfo.DomainShortName
	WriteLog("Found domain '" & sDomain & "'")

	Dim objSysInfo, objUser

	Set objSysInfo = CreateObject("ADSystemInfo")
	Set objUser = GetObject("LDAP://" & objSysInfo.UserName)
	writelog("Now attempting to set the comment field for user object " & objSysInfo.Username _
	& " with text '" & sComputer & "'")

	objUser.Comment = sComputer
	objUser.SetInfo
	If err.number <> 0 then
		Call LogError(err.number,err.description,"'SetAttribute'")
	Else
		Call WriteLog("Comment attribute updated.")
	End If

	err.clear
End Sub 'SetAttribute

'------------------------------------------------------------------------------'
' Subroutine LogError - this logs any errors to the logfile
'
' Correct input is LogError("5","This file does not exist","RemoveNetworkDrives")
'##############################################################################'
Sub LogError(eNumber,eDescription,sMsg)
	on error resume next
	If warn=1 then msgbox "Error " & eNumber & " - " & replace(eDescription,vbnewline,"") & " This error occured " _
	& "while calling " & sMsg
	log.writeline("[" & Now & "] - Error " & eNumber & " - " & replace(eDescription,vbnewline,"") & " This error occured " _
		& "while calling " & sMsg)
End Sub 'LogError

'------------------------------------------------------------------------------'
' Subroutine WriteLog - this logs the logon process to a .log file
'
' Correct input is WriteLog("text")
'##############################################################################'
Sub WriteLog(sMsg)

on error resume next
	'wscript.echo "[" & Now & "] - " & sMsg
	log.writeline("[" & Now & "] - " & sMsg)
	'If right(lcase(wscript.fullname),11) = "cscript.exe" then wscript.echo sMsg
End Sub 'WriteLog

'------------------------------------------------------------------------------'
' Subroutine VerifyDrive - this checks to see if a passed letter and path
' exist on a computer.
'
' Correct input is WriteLog("r:","\\server\share")
'##############################################################################'
Function VerifyDrive(sLetter,sPath)
	on error resume next
	Set oShell = CreateObject("Shell.Application")
	sLetter = replace(sLetter,":","")
	sName = oShell.NameSpace(sLetter & ":\").Self.Name
	bMapped = false

	aTemp = split(sPath,"\")
	aTemp1 = split(sName," on ")
	aTemp2 = split(replace(replace(replace(aTemp1(1),"'","|"),"(","|"),")","|"),"|")

	for i = 0 to ubound(aTemp2)
		if lcase(aTemp2(i)) = lcase(aTemp(0)) then bServerMatch = true
		if lcase(aTemp1(i)) = lcase(aTemp(1)) then bShareMatch = true
	next

	If bServerMatch = true and bShareMatch = true then bMapped = true

	VerifyDrive = bMapped
End Function

'------------------------------------------------------------------------------'
' Subroutine RemoveNetworkDrives - unmaps specified drives (if bRemoveDrives
' is set to 'true')
'
' Correct input is RemoveNetworkDrives("X")
' where 'X' is the drive letter you wish to disconnect.
'##############################################################################'
Sub RemoveNetworkDrives(sDriveLetter,sText)
	on error resume next

	WriteLog("Attempting to remove drive letter '" & sDriveLetter & "' " & sText)

	objNetwork.RemoveNetworkDrive sDriveLetter,TRUE, TRUE

	If err.number <> 0 then
		If err.number = -2147022492 Then
			Call WriteLog("Either this drive is locked by the system, or the resource does not exist.")
		ElseIf err.number = -2147022646 Then
			Call WriteLog("This network drive (" & sDriveLetter & "), does not exist.")
		Else
			Call LogError(err.number,err.description,"'" & "RemoveNetworkDrives(" & sDriveLetter & ")'")
		End If
	Else
		Call WriteLog("Removal successful.")
	End If
	err.clear
End Sub 'RemoveNetworkDrives

'------------------------------------------------------------------------------'
' Subroutine RenameDrives - Renames network drives if specified.
'
' Correct input is RenameDrives("x:\\server\share")
' where 'x' is the drive letter you wish to rename, \\server\share is the
' servername and share associated with the letter.
'##############################################################################'
Sub RenameDrives(sPath)
	Set oShell = CreateObject("Shell.Application")
	'Split the given path, using "\" as a delimiter.
	aDriveDesc = split(sPath,"\")
	Call WriteLog("Attempting to rename " & aDriveDesc(0) & " to [" & aDriveDesc(ubound(aDriveDesc)) & " on " & aDriveDesc(2) & "]")
	oShell.NameSpace(aDriveDesc(0) & "\").Self.Name = "[" & aDriveDesc(ubound(aDriveDesc)) & " on " & aDriveDesc(2) & "]"
	wscript.sleep 1000
	If err.number <> 0 then
		call LogError(err.number,err.description,"'RenameDrives" _
		& "(" & sPath & ")")
	Else
		Call WriteLog("Rename succeeded.")
	End If
	err.clear
End Sub 'RenameDrives

'------------------------------------------------------------------------------'
' Function inOU - this returns the current OU of either a user or computer
'
' Correct input is
' inOU("OU=domain users,OU=Departments",user|computer,true|false)
'
' bVerbose tells the function whether or not to write to the logfile
'##############################################################################'
Function inOU(strOUValue,sType,bVerbose)

	If lcase(sType) = "user" then
		tempOU = user_OU
	ElseIf lcase(sType) = "computer" then
		tempOU = computer_OU
	End If

	If instr(lcase(tempOU),strOUValue) then
		binOU = true
		sText = " is "
	Else
		binOU = false
		sText = " is not "
	End If

	If bVerbose = true then WriteLog(sType & " object" & sText & "a member of " & strOUValue)
	inOU = binOU
End Function

'------------------------------------------------------------------------------'
' Subroutine AddNetworkPrinter - Add a network printer if specified.
'
' Correct input is AddNetworkPrinter("\\server\share",false)
' ...if you want to set the printer as default, then set the second argument
' to 'true'
'##############################################################################'
Function AddNetworkPrinter(sPrinter,bSetPrinterAsDefault)
	on error resume next
	Call WriteLog("Attempting to add printer: " & sPrinter)
	'Wscript.echo "Installing printer " & sPrinter

	objNetwork.AddWindowsPrinterConnection sPrinter
	'wscript.echo err.number & " " & err.description

	If bSetPrinterAsDefault = true then objNetwork.SetDefaultPrinter sPrinter

	If err.number <> 0 then
		If err.number = -2147023095 Then
			Call WriteLog("No printer '" & sPrinter & "' was found on the specified server.")
		Else
			Call LogError(err.number,err.description,"'AddNetworkPrinter" _
			& "(" & sPrinter & "," & bSetPrinterAsDefault & ")")
		End If
	Else
		Call WriteLog("Printer add succeeded.")
	End If

	err.clear
End Function 'AddNetworkPrinter

'------------------------------------------------------------------------------'
' Function RemoveNetworkPrinter - this removes a printer specified by the
' argument 'sprinter'
'
' Correct input is RemoveNetworkPrinter("\\server\printershare")
'
'##############################################################################'
Function RemoveNetworkPrinter(sPrinter)

	on error resume next

	Call WriteLog("Attempting to remove printer: " & sPrinter)

	objNetwork.RemovePrinterConnection sPrinter

	If err.number <> 0 then
		If err.number = -2147022646 then
			WriteLog(sPrinter & " is not installed locally, skipping removal")
		Else
			Call LogError(err.number,err.description,"'AddNetworkPrinter" _
			& "(" & sPrinter & "," & bSetPrinterAsDefault & ")")
		End If
	Else
		Call WriteLog("Printer remove succeeded.")
	End If

End Function

'------------------------------------------------------------------------------'
' Subroutine MapNetworkDrives - maps specified network drives.
'
' Correct input is MapNetworkDrives("X:","\\server\share)
' where 'X' is the drive letter you wish to map.
'##############################################################################'
Sub MapNetworkDrives(sDriveLetter,sPath)
	Set oShell = CreateObject("Shell.Application")
	sDriveLetter = replace(sDriveLetter,":","")
	on error resume next
	'Map network drive with the drive letter and path

	'Call WriteLog("Attempting to map drive " & sDriveLetter & ": to " & sPath)
	'Wscript.echo "Mapping drive " & sDriveLetter & ": to " & sPath
	objNetwork.MapNetworkDrive sDriveLetter & ":" , sPath, true

	If err.number <> 0 Then
		If err.number = -2147024811 then
			Call WriteLog(sDriveLetter & ":\ is already mapped as " & oShell.NameSpace(sDriveLetter & ":\").Self.Name)
		Else
			Call LogError(err.number,err.description,"'MapNetworkDrives" & "(" & sDriveLetter & "," & sPath & ")")
		End If
	Else
		Call WriteLog(sDriveLetter & ":\ drive map succeeded to " & sPath)
	End If
	'Set objNetwork = nothing
	err.clear
End Sub 'MapNetworkDrives

'------------------------------------------------------------------------------'
' Subroutine RunProcess - Runs a specified executable or script.
'
' Correct input is RunProcess("c:\process.exe",false)
'
' If you want the process to wait until completion before going on, set the
' second argument to 'true'
'##############################################################################'
Sub RunProcess(sProcessName,blnWait)
	Call WriteLog("Attempting to run process: " & sProcessName)
	on error resume next
	sResult = WshShell.Run(sProcessName,1,blnWait)
	If err.number <> 0 then
		If err.number = -2147024894 then
			Call WriteLog("Could not locate " & sProcessName)
		Else
			Call LogError(err.number,err.description,"'RunProcess" _
			& "(" & sProcessName & "," & blnWait & ")")
		End If
	Else
		If sResult = 0 then sMsg = "The process started successfully."
		Call WriteLog(sMsg)
	End if
	err.clear
End Sub 'RunProcess

Function GetIPAddress
	on error resume next
	strcomputer = "."
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	Set colItems = objWMIService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

	for each objitem in colitems

	sIPAddress = Join(objitem.IPAddress, " - ")

	sIPAddress1 = sIPAddress & delim & sIPAddress1
	delim = "|"
	'msgbox sIPAddress
	next

	GetIPAddress = sIPAddress1
End Function 'GetIPAddress

'------------------------------------------------------------------------------'
' Subroutine AddSupportGroup - adds group or user to local group
'
' Correct input is AddSupportGroup("monogram\rockford mis3","administrators",false)
' or AddSupportGroup("username","administrators",false)
'
' The group/user name must be a member of the domain that the logged on user
' is in order for this process to work properly. Set the third argument to
' 'true' if you want this to process on servers as well.
'
' Of course, the logged on user must have administrative privileges on the
' computer in order to add users to the administrative group.
'
' A separate startup script could be generated to get around this (runs at
' computer startup rather than user logon...
'##############################################################################'
Sub AddSupportGroup(sObject,sLocalGroup,bDontRunOnServers)

	Set objWMIService = GetObject("winmgmts:\\" & sComputer & "\root\cimv2")
	Set colOperatingSystems = objWMIService.ExecQuery("Select Caption,ServicePackMajorVersion from Win32_OperatingSystem")

	If err.number <> 0 then call LogError(err.number,err.description,"'AddSupportGroup" _
	& "(" & sObject & "," & sLocalGroup & "," & bDontRunOnServers & ")")

	'Split the domain\username into two parts...
	aName = split(sObject,"\")

	'If no domain is found, then exit the sub.
	If aName(0) = "" then
		WriteLog("You must specify a domain and a backslash before the username when calling the 'AddSupportGroup' Subroutine.")
		Exit Sub
	Else
		'Get the domain name, place it in 'sDomain' variable.
		sDomain = aName(0)
		'Get the username from the domain\username combo, place it in the 'sName'
		' variable.
		sName = aName(1)

	End If

	For Each objOperatingSystem in colOperatingSystems

		If objOperatingSystem.ServicePackMajorVersion <> "" then
			sSPVersion = "SP" & objOperatingSystem.ServicePackMajorVersion
		Else
			sSPVersion = ""
		End if

		Call WriteLog("Found '" & objOperatingSystem.caption & " " & sSPVersion & "' as the installed " _
		& "operating system for " & sComputer)

		If err.number <> 0 then

			'If there is an error finding the Operating system version, then quit to err
			' on the side of caution.
			Exit Sub
		End If

		'If we've specified not to run on servers, then quit the script.
		If bDontRunOnServers = true Then
			If InStr(objOperatingSystem.Caption, "Server") or InStr(objOperatingSystem.Caption, "Powered") Then
					If warn = 1 Then MsgBox "This script is set to run on workstations only. Please modify the '" _
					& sLocalGroup & "' group manually." & vbcrlf & vbcrlf & "Now quitting.",48,"Not for use on servers!"
					Exit Sub
			End If
		End If
	Next
	on error resume next

	Call writelog("Attempting to add " & sdomain & "\" & sName & " to " & sLocalGroup & " on " & sComputer)
	'Attach to the group object using the WinNT provider.
	Set oGroup = GetObject("WinNT://"& sComputer &"/" & sLocalGroup)

	'Add the specified domain\user to the local group that we referenced above.

	oGroup.Add "WinNT://" & sDomain & "/" & sName

	If err.Number <> 0 Then
		If err.number = -2147023518 Then
			Call WriteLog("This object is already a member of the " & sLocalGroup & " group on this computer")
		Else
			Call LogError(err.number,err.description,"'AddSupportGroup" _
			& "(" & sObject & "," & sLocalGroup & "," & bDontRunOnServers & ")")
		End If
	Else
		Call WriteLog("Object '" & sName & "' added successfully to group " & sLocalGroup)
	End If

	on error resume next
	err.clear
End Sub 'AddSupportGroup

'------------------------------------------------------------------------------'
'Function IsMember - This will check the domain and logged on user ID to find
' out whether or not it is a member of the specified group in 'strGroup'
' the title must match the group name exactly.
'
' Correct input is IsMember("rockford mis3")
' The group must be a member of the same domain as the user that is logging
' in.
'
' Now supports one level of recursion - user may be in a group that is a member of the target group.
'##############################################################################'
Public Function IsMember(groupName,bVerbose)

	Dim flagIsMember
	flagIsMember = False

	For Each memberOf In objGroupList
		If StrComp(groupName, memberOf, 1) = 0 Then
			flagIsMember = True
			Exit For
		End If
	Next

	IsMember = flagIsMember
	If flagIsMember Then
		WriteLog("Member of " & groupName)
	Else
		WriteLog("Not member of " & groupName)
	End If

End Function 'IsMember

'*******************************************************************************
'Function ReportFileStatus(filespec)
'Determines the existence of a file - reports 'true' or 'false' depending on
' the results.
'*******************************************************************************
Function ReportFileStatus(filespec)
	on error resume next
	If (objfso.FileExists(filespec)) Then
		msg = "true"
	Else
		msg = "false"
	End If
	if err.number <> 0 then call logerror(err.number,err.description,"'ReportFileStatus(" & filespec & ")")
	ReportFileStatus = msg
End Function 'ReportFileStatus

Function SortDictionary(objDict,intSort)
	' declare our variables
	Dim strDict()
	Dim objKey
	Dim strKey,strItem
	Dim X,Y,Z

	' get the dictionary count
	Z = objDict.Count

	' we need more than one item to warrant sorting
	If Z > 1 Then
		' create an array to store dictionary information
		ReDim strDict(Z,2)
		X = 0
		' populate the string array
		For Each objKey In objDict
			strDict(X,dictKey) = CStr(objKey)
			strDict(X,dictItem) = CStr(objDict(objKey))
			X = X + 1
		Next

		' perform a a shell sort of the string array
		For X = 0 to (Z - 2)
			For Y = X to (Z - 1)
				If StrComp(strDict(X,intSort),strDict(Y,intSort),vbTextCompare) > 0 Then
					strKey = strDict(X,dictKey)
					strItem = strDict(X,dictItem)
					strDict(X,dictKey) = strDict(Y,dictKey)
					strDict(X,dictItem) = strDict(Y,dictItem)
					strDict(Y,dictKey) = strKey
					strDict(Y,dictItem) = strItem
				End If
			Next
		Next

	' erase the contents of the dictionary object
	objDict.RemoveAll

	' repopulate the dictionary with the sorted information
	For X = 0 to (Z - 1)
		objDict.Add strDict(X,dictKey), strDict(X,dictItem)
	Next

	End If

End Function

Function EnumGroups(ByVal objADObject)
' Recursive subroutine to enumerate user group memberships.
' Slightly Modified Version of script originally created and posted
' at http://www.rlmueller.net/Programs/EnumUserGroups.txt
' Includes nested group memberships.
	Dim colstrGroups, objGroup, j
	objGroupList.CompareMode = vbTextCompare
	colstrGroups = objADObject.memberOf

	If (IsEmpty(colstrGroups) = True) Then
		Exit Function
	End If
	If (TypeName(colstrGroups) = "String") Then
		' Escape any forward slash characters, "/", with the backslash
		' escape character. All other characters that should be escaped are.
		colstrGroups = Replace(colstrGroups, "/", "\/")
		Set objGroup = GetObject("LDAP://" & colstrGroups)
		If (objGroupList.Exists(objGroup.sAMAccountName) = False) Then
			objGroupList.Add objGroup.sAMAccountName, True
			Call EnumGroups(objGroup)
		End If
		Exit Function
	End If
	For j = 0 To UBound(colstrGroups)
		' Escape any forward slash characters, "/", with the backslash
		' escape character. All other characters that should be escaped are.
		colstrGroups(j) = Replace(colstrGroups(j), "/", "\/")
		Set objGroup = GetObject("LDAP://" & colstrGroups(j))
		If (objGroupList.Exists(objGroup.sAMAccountName) = False) Then
			objGroupList.Add objGroup.sAMAccountName, True
			Call EnumGroups(objGroup)
		End If
	Next
End Function

'''' End Extra Large Logon Script
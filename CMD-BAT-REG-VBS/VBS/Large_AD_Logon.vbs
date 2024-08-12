'AD Logon Script-----------------------------------------------------------------------------'

' Standard domain login script
' Maps Network Drives & Printers based on group membership.
' -----------------------------------------------------------------'
' ***NOTES***
' Using VBS to perform an Active Directory LDAP query of current user's group memberships does not enumerate the primary group.
'	So you cannot reference a user's primary group to perform actions.
' Furthormore, all users must be a direct member of the group you are referencing to perform actions.
'	Since VBS can only enumerate the direct members of a group, if a group contains another group which contains the users
'	the script will not be able to see the users in the nested group.

' If you expect a Batch File should be running, but it is not, please check the OnOff & PreFlightCheck values in the array.
' The OnOff value must be set to "on" for a batch file to be run.
' Also, if the PreFlightCheck value is set to "yes" then you must have a valid PreFlightCode **AND** you must insert your
'	pre-flight parameters into the script.
' The PreFlightCheck = "yes" tells the script that there are pre-requisites that must be met before the batch file can be run.
' Even if the PreFlightCheck is set to NO, it is recommended that you set the PreFlightCode for troubleshooting purposes.

' There is a strMessage and a MsgBox that can be uncommented out in the runBatch Function, if you suspect a problem.
' This has been added for troubleshooting purposes because login scripts can grow to be quite large.
' -----------------------------------------------------------------'
Option Explicit

' Create the objects that will be used in the script.
Dim objNetwork, objFSO, WshShell 
Set objNetwork = CreateObject("WScript.Network")
Set objFSO = CreateObject("scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")

' Get all groups the current user is a member of and join them together in CSV:
Dim objUser, CurrentUser, strGroup, strFunctionName
Set objUser = CreateObject("ADSystemInfo")
Set CurrentUser = GetObject("LDAP://" & objUser.UserName)
strGroup = LCase(Join(CurrentUser.MemberOf))
' -----------------------------------------------------------------'
' Build array tables with reference information:
' -----------------------------------------------------------------'
' DL = The drive letter and paths of all mappable drives - sorted by drive letter.
' PTR = The port names in the windows registry & paths of all shared network printers.
' BF = A list of batch files for running various commands like updating the registry or uninstalling windows components.
' GRPS = Listing of all groups used in the process of mapping drives & printers
'	(GRPS must always be the last array enumerated because it contains nested array info)
' -----------------------------------------------------------------'
Dim DL(14) 'Drive Letters
'			Drive(0)		Path(1)
DL(0) = array("H", 		"\\domain.com\root\netapps")
DL(1) = array("I", 		"\\domain.com\root\implementation")
DL(2) = array("J", 		"\\domain.com\root\accounting")
DL(3) = array("K", 		"\\domain.com\root\itdepartment")
DL(4) = array("L", 		"\\domain.com\root\tap")
DL(5) = array("M", 		"\\domain.com\root\tas")
DL(6) = array("N", 		"\\domain.com\root\ecm")
DL(7) = array("O", 		"\\domain.com\root\operations")
DL(8) = array("P",	 	"\\domain.com\root\iso")
DL(9) = array("U", 		"\\domain.com\root\clientdata")
DL(10) = array("R", 	"\\rum\co$")
DL(11) = array("S", 	"\\domain.com\root\ukt")
DL(12) = array("T", 	"\\domain.com\root\teleform")
DL(13) = array("W", 	"\\domain.com\root\marketing")
DL(14) = array("X", 	"\\domain.com\root\scanner")
'
Dim PTR(3) 'Printers
'					Port(0)					Path(1)										Name(2)
PTR(0) = array("IP_10.1.1.93",				"\\domaindc02\XEROX",						"XEROX")
PTR(1) = array("IP_10.1.1.98,IP_10.1.1.99",	"\\domaindc02\HP8150Printroom",				"HP8150Printroom")
PTR(2) = array("IP_10.1.1.112",				"\\domaindc02\HP4100Acctng",				"HP4100Acctng")
PTR(3) = array("IP_10.1.1.97",				"\\domaindc02\RICOH",						"RICOH")
'
Dim BF(0) 'Batch Files to Run (exe, bat, txt..)
'				BatchFile(0)						Description(1)				OnOff(2) PreFlightCheck?(3) PreFlightCode(4)
BF(0) = array("\\domaindc01\netlogon\nogames.bat", "uninstall Microsoft Windows Games", "on", "yes", "nogames")
'
Dim GRPS(22) 'Active Directory Groups
'					GroupName(0)			MapDrive(1)			InstallPrinter(2)	SetDefaultPrinter(3)	BatchFiles(4)
GRPS(0) = array("cn=accountingdata",		array(DL(2)),				"",				"",					"")
GRPS(1) = array("cn=clientdata", 			array(DL(9)),				"",				"",					"")
GRPS(2) = array("cn=computerops",			array(DL(10)),				"",				"",					"")
GRPS(3) = array("cn=ecm",					array(DL(6)),				"",				"",					"")
GRPS(4) = array("cn=implementation",		array(DL(1)),				"",				"",					"")
GRPS(5) = array("cn=iso",	 				array(DL(8)),				"",				"",					"")
GRPS(6) = array("cn=itdepartment",			array(DL(3)),				"",				"",					"")
GRPS(7) = array("cn=marketing", 			array(DL(13)),				"",				"",					"")
GRPS(8) = array("cn=netapps",	 			array(DL(0)),				"",				"",					"")
GRPS(9) = array("cn=operations", 			array(DL(7)),				"",				"",					"")
GRPS(10) = array("cn=scanner",	 			array(DL(14)),				"",				"",					"")
GRPS(11) = array("cn=tapdata", 				array(DL(4)),				"",				"",					"")
GRPS(12) = array("cn=tasdata", 				array(DL(5)),				"",				"",					"")
GRPS(13) = array("cn=teleformdata", 		array(DL(12)),				"",				"",					"")
GRPS(14) = array("cn=ukt",		 			array(DL(11)),				"",				"",					"")

GRPS(15) = array("cn=ptr_hp4100act",		"",							array(PTR(2)),	"",					"")
GRPS(16) = array("cn=ptr_hp8150pr",		 	"",							array(PTR(1)),	"",					"")
GRPS(17) = array("cn=ptr_ricoh",			"",							array(PTR(3)),	"",					"")
GRPS(18) = array("cn=ptr_xerox",			"",							array(PTR(0)),	"",					"")
GRPS(19) = array("cn=ptrdefault_hp4100act",	"",							"",				array(PTR(2)),		"")
GRPS(20) = array("cn=ptrdefault_ricoh",		"",							"",				array(PTR(3)),		"")
GRPS(21) = array("cn=ptrdefault_hp8150pr",	"",							"",				array(PTR(1)),		"")

GRPS(22) = array("cn=uninstallgames", 		"",							"",				"",					array(BF(0)))


' Define the variables that will be used in the logic loop
Dim RowGRPS, ColGRPS, RowNEST, ColNest
' RowGRPS = the row number of the GRPS array
' ColGRPS = the column number of the GRPS array
' RowNEST = the row number of the NESTED array
' ColNEST = the row number of the NESTED array
' -----------------------------------------------------------------'
' Begin the looping process to check current group memberships
' Take action accordingly (map drives, printers, etc.)
' -----------------------------------------------------------------'
' sets RowGRPS = 0 of the GRPS array and increments to the Upper Boundry for each loop (all Rows in the GRPS array)
For RowGRPS = 0 to Ubound(GRPS)
	WScript.echo "Checking: " & GRPS(RowGRPS)(0)
	' Loop through each group name in the array and check to see if you are a member.  GRPS(RowGRPS)(0) is always Group Name, so:
	If InStr(strGroup, lcase(GRPS(RowGRPS)(0))) Then
		WScript.echo " Yes, you are a member of " &GRPS(RowGRPS)(0)
		' If you are a member, get the drive letter(s) the script will map.  Since GRPS(RowGRPS)(1) is the ALWAYS the nested DL(n) reference:
		ColGRPS = 1
		If IsArray(GRPS(RowGRPS)(ColGRPS)) Then
			' sets RowNEST = 0 of the DL array and increments to the Upper Boundry for each loop (all Rows in the DL array)
			For RowNEST = 0 to Ubound(GRPS(RowGRPS)(ColGRPS))
				If isArray(GRPS(RowGRPS)(ColGRPS)(RowNEST)) Then
					' we need to check ONLY the drive letter against currently mapped drive letters.  Since GRPS(RowGRPS)(ColGRPS)(RowNEST)(0) is ALWAYS drive letter:
					ColNEST =0
					WScript.echo "  Script will map " & GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST) & ": to " &GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST+1)
						' Compare group drives to currently mapped drives, unmap if already exists
						If (doesDriveExist(GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST)& ":")) Then
							WScript.echo "	Drive " & GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST)& ": already exists!  Removing..."
							objNetwork.RemoveNetworkDrive GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST) & ":"
						End If
					' Map the drive
					WScript.echo "	Mapping drive " & GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST) & ":..."
					objNetwork.MapNetworkDrive GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST) & ":", GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST+1)
				End If
			Next
		End if
		ColGRPS = 2
		If IsArray(GRPS(RowGRPS)(ColGRPS)) Then
			' sets RowNEST = 0 of the PTR array and increments to the Upper Boundry for each loop (all Rows in the PTR array)
			For RowNEST = 0 to Ubound(GRPS(RowGRPS)(ColGRPS))
				If isArray(GRPS(RowGRPS)(ColGRPS)(RowNEST)) Then
					' we need to check ONLY the Port against currently mapped printers.  Since GRPS(RowGRPS)(ColGRPS)(RowNEST)(0) is ALWAYS port:
					ColNEST = 0
					WScript.echo "  Script will connect to printer " & GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST+2)
						' Compare group drives to currently connected printers, disconnect if already exists
						If (doesPtrExist(GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST))) Then
							WScript.echo "Printer " & GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST+2)& " already exists!  Removing..."
							objNetwork.RemovePrinterConnection GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST+1)
						End If
					' Connect to the printer and set it as default
					WScript.echo "  Connecting to printer " & GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST+2) & "..."
					objNetwork.AddWindowsPrinterConnection GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST+1)
				End If
			Next
		End if
		ColGRPS = 3
		If IsArray(GRPS(RowGRPS)(ColGRPS)) Then
			' sets RowNEST = 0 of the PTR array and increments to the Upper Boundry for each loop (all Rows in the PTR array)
			For RowNEST = 0 to Ubound(GRPS(RowGRPS)(ColGRPS))
				If isArray(GRPS(RowGRPS)(ColGRPS)(RowNEST)) Then
					' we need to check ONLY the Port against currently mapped printers.  Since GRPS(RowGRPS)(ColGRPS)(RowNEST)(0) is ALWAYS port:
					ColNEST = 0
					WScript.echo "  Based on group membership, your default printer is " & GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST+2) & "..."
					objNetwork.SetDefaultPrinter GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST+1)
				End if
			Next
		End if
		ColGRPS = 4
		If IsArray(GRPS(RowGRPS)(ColGRPS)) Then
			' sets RowNEST = 0 of the BF array and increments to the Upper Boundry for each loop (all Rows in the BF array)
			For RowNEST = 0 to Ubound(GRPS(RowGRPS)(ColGRPS))
				If isArray(GRPS(RowGRPS)(ColGRPS)(RowNEST)) Then
					' Since GRPS(RowGRPS)(ColGRPS)(RowNEST)(0) is ALWAYS the batch file to run:
					ColNEST = 0
					WScript.echo "  Checking to see if " & GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST) & " will run..."
					If (runBatch((GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST+2)),(GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST+3)),(GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST+4)))) Then
						WScript.echo "	Yes, " & GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST) & " will be run."
						WshShell.run (GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST)),0,False
					Else
						WScript.echo "	No, " & GRPS(RowGRPS)(ColGRPS)(RowNEST)(ColNEST) & " will NOT be run."
					End If
				End If
			Next
		End if
	Else
		WScript.echo " No, you are not a member of " &GRPS(RowGRPS)(0)
	End If
Next
WScript.quit


Function doesDriveExist(driveLetter)
	Dim objNetwork, enumDrives, RowGRPS
	
	Set objNetwork = CreateObject("WScript.Network")
	Set enumDrives = objNetwork.EnumNetworkDrives
	
	doesDriveExist = false
	
	For RowGRPS = 0 to enumDrives.Count - 1 Step 2
		If enumDrives.Item(RowGRPS) = driveLetter Then
			doesDriveExist = true
			Exit Function
		End If
	Next
End Function


Function doesPtrExist(ptrPort)
	Dim objNetwork, enumPrinters, RowGRPS
	
	Set objNetwork = CreateObject("WScript.Network")
	Set enumPrinters = objNetwork.EnumPrinterConnections
	
	doesPtrExist = false
	
	For RowGRPS = 0 to enumPrinters.Count - 1 Step 2
		If enumPrinters.Item(RowGRPS) = ptrPort Then
			doesPtrExist = true
			Exit Function
		End If
	Next
End Function


Function runBatch(x,y,z)
	Dim objFSO, WshShell, strMessage
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set WshShell = WScript.CreateObject("WScript.Shell")
	
	runBatch = false
	
	If x = "on" Then
		If y = "yes" Then
			If z = "nogames" Then
				If objFSO.FileExists("c:\windows\system32\freecell.exe") Then
					runBatch = true
					Exit Function
				End If
			Else
				runBatch = false
				'strMessage = "PreFlightCheck FAILED!  PreFlightCheck value set to YES, but PreFlightCode [" & z & "] was not found!  The batch file will NOT be run."
				'MsgBox strMessage, vbOKOnly + vbExclamation, "PreFlightCode not found!"
				Exit Function
			End If
		Else
		runBatch = true
		'strMessage = "PreFlightCheck was not required!  PreFlightCheck value is set to NO.  The batch file [" & z & "] will be run without any pre-requisites."
		'MsgBox strMessage, vbOKOnly + vbExclamation, "PreFlightCheck not Required!"
		Exit Function
		End If
	End If
End Function

'''End AD Logon-------------------------------
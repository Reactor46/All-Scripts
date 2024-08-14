'Title:        Uninstall Old Citrix Clients
'Date:         5/25/2010
'Author:       Gregory Strike
'
'Purpose:      The following script will remove any Citrix clients found
'              but leave the latest Citrix Online Plug-in v12.0.  The version
'              numbers can be modified below for later releases of the Citrix
'              online plug-in.
'
'              This script was originally written as an Active Directory Startup
'              script to assist in deploying the latest Citrix Online Plug-in.
'              It can, however, be ran manually using: CScript.exe [ScriptName.vbs]
'
'Requirements: Administrative Privileges
const HKEY_LOCAL_MACHINE = &H80000002
 
strComputer = "."
 
Function UninstallApp(strDisplayName, strVersion, strID, strUninstall)
	Dim objShell
	Dim objFS 
	'WScript.Echo "Attempting to uninstall: " & strDisplayName & " v" & strVersion
 
	If strID = "" Then  'We don't know the GUID of the app
		'Look at the Uninstall string and determine what is the
		'executable and what are the command line arguments.
		Set objFS = CreateObject("Scripting.FileSystemObject")
 
		strExecutable = ""
 
		'Start from the beginning of the string and see if we can fine the excutable in the string
		For X = 0 to Len(strUninstall)
			strExecutableTest = Left(strUninstall, X)
			strExecutableTest = Replace(strExecutableTest, """", "")
			'Test to see if the current string is a file.
			If objFS.FileExists(strExecutableTest) Then
				strExecutable = Trim(strExecutableTest)
				intExecLength = X
			End If
		Next
 
		If strExecutable = "" Then
			'WScript.Echo "Bad string or the executable does not exist: " & strUninstall
			Exit Function
		Else
			strArguments = Right(strUninstall, Len(strUninstall) - intExecLength)
			'WScript.Echo "The executable is: " & strExecutable
			'WScript.Echo "The arguments are: " & strArguments
		End If
 
		Uninstall = """" & strExecutable & """ " & strArguments 
 
		If InStr(Uninstall, "ISUNINST.EXE") > 0 Then
			Uninstall = Uninstall & " -a"
		End If
	Else 'We have the GUID
		Uninstall = """MSIEXEC.EXE"" /PASSIVE /X" & strID
	End If
 
	'WScript.Echo "...Executing: " & Uninstall
 
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.Run Uninstall, 1 , 1
 
	Set objShell = Nothing
End Function
 
'WScript.Echo ""
'WScript.Echo Now() & " - Searching for old Citrix Clients..."
 
Set ObjWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
ObjWMI.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
 
For Each Product In arrSubKeys
	objWMI.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath & "\" & Product, "DisplayName", strDisplayName
	objWMI.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath & "\" & Product, "DisplayVersion", strVersion
	objWMI.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath & "\" & Product, "UninstallString", strUninstall
 
	strName = UCase(strDisplayName)
 
	'Grab the GUID of the MSI if available.
	If Left(Product, 1) = "{" And Right(Product, 1) = "}" Then
		strID = Product
	Else
		strID = ""
	End If
 
	'Determine version of the Product
	If strVersion <> "" Then
		VersionArray = Split(strVersion, ".")
		If UBound(VersionArray) > 0 Then
			'Verify that only numbers are in the version string
			If IsNumeric(VersionArray(0)) And IsNumeric(VersionArray(1)) Then
				Version = CDbl(VersionArray(0) & "." & VersionArray(1))	
			Else
				Version = ""
			End If
		End If
	Else
		Version = ""
	End If
 
	'Citrix has used many different Client names throughout the years.  This
	'should capture most, if not all, of them.
	
	If strName = "CITRIX ICA CLIENT" Then
		UninstallApp strDisplayName, strVersion, strID, strUninstall
	End If
 
	If strName = "CITRIX PROGRAM NEIGHBORHOOD" Then
		UninstallApp strDisplayName, strVersion, strID, strUninstall
	End If
 
	If strName = "METAFRAME PRESENTATION SERVER CLIENT" Then
		UninstallApp strDisplayName, strVersion, strID, strUninstall
	End If
 
	If strName = "CITRIX PRESENTATION SERVER CLIENT" Then
		UninstallApp strDisplayName, strVersion, strID, strUninstall
	End If
 
	'Added
	If strName = "CITRIX ICA WEB CLIENT" Then
		'Unfortunately, CTXSETUP.exe doesn't allow a silent uninstall.
		UninstallApp strDisplayName, strVersion, strID, strUninstall
	End If	
 
	If Left(strName, 21) = "CITRIX ONLINE PLUG-IN" Then
		'Comment the IF & End If lines out if you want to remove all "Citrix Online Plug-in" 's.
		'This If statement will leave version 12.0 and higher installed.
		If Version < 12.0 Then
		UninstallApp strDisplayName, strVersion, strID, strUninstall
		End If
	End If
Next
 
'WScript.Echo Now() & " - Search Complete."
'WScript.Echo ""
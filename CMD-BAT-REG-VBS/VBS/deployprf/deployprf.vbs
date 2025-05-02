'This script determines if a specified mail profile already exists.
'If it doesn't, it will set the path to the prf-file containing
'the mail profile configuration settings.

'Script created by: Robert Sparnaaij
'For more information about this file see;
'http://www.howto-outlook.com/howto/deployprf.htm


'=====BEGIN EDITING=====

'Name of mail profile as in the prf-file
ProfileName = "Outlook"

'Path to the prf-file
ProfilePath = "\\server\share\profile.prf"

'Increase the ProfileVersion whenever you want to reapply the prf-file
ProfileVersion = 1

'======STOP EDITING UNLESS YOU KNOW WHAT YOU ARE DOING=====

const HKEY_CURRENT_USER = &H80000001
const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
Set objRegistry = GetObject("winmgmts:\\" & _
    strComputer & "\root\default:StdRegProv")
 
strKeyProfilePath = _ 
	"Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\" _ 
	& ProfileName & "\9375CFF0413111d3B88A00104B2A6676"
strLastChangeVer = "LastChangeVer"
objRegistry.GetBinaryValue _
	HKEY_CURRENT_USER,strKeyProfilePath,strLastChangeVer,strValueLastChangeVer

If ProfileVersion > 1 Then
    strKeyProfileVersionPath = "SOFTWARE\HowTo-Outlook\DeployPRF"
    strProfileVersionName = ProfileName
    objRegistry.GetDWORDValue _
    	HKEY_CURRENT_USER,strKeyProfileVersionPath,strProfileVersionName,strValueProfileVersion

    If IsNull(strValueProfileVersion) OR ProfileVersion > strValueProfileVersion Then
	ReapplyPrf = True
    End If
End If

If IsNull(strValueLastChangeVer) OR ReapplyPrf Then
    'The mail profile doesn't exist yet so we'll set the the ImportPRF key and remove the FirstRun keys

    'Determine path to outlook.exe
    strKeyOutlookAppPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"
    strOutlookPath = "Path"
    objRegistry.GetStringValue _
    	HKEY_LOCAL_MACHINE,strKeyOutlookAppPath,strOutlookPath,strOutlookPathValue

    'Verify that the outlook.exe and the configured prf-file exist
    Set objFSO = CreateObject("Scripting.FileSystemObject") 
    If objFSO.FileExists(strOutlookPathValue & "outlook.exe") AND objFSO.FileExists(ProfilePath) Then

	'Determine version of Outlook
	strOutlookVersionNumber = objFSO.GetFileVersion(strOutlookPathValue & "outlook.exe")
	strOutlookVersion = Left(strOutlookVersionNumber, inStr(strOutlookVersionNumber, ".0") + 1)

	'Create the Setup key, set the ImportPRF value and delete the First-Run values.
	strKeyOutlookSetupPath = "SOFTWARE\Microsoft\Office\" & strOutlookVersion & "\Outlook\Setup"

	strImportPRFValueName = "ImportPRF"
	strImportPRFValue = ProfilePath
	objRegistry.CreateKey HKEY_CURRENT_USER,strKeyOutlookSetupPath
	objRegistry.SetStringValue HKEY_CURRENT_USER,_
 	    strKeyOutlookSetupPath,strImportPRFValueName,strImportPRFValue

	strFirstRunValueName = "FirstRun"
	objRegistry.DeleteValue HKEY_CURRENT_USER,_
 	    strKeyOutlookSetupPath,strFirstRunValueName

	strFirstRun2ValueName = "First-Run"
	objRegistry.DeleteValue HKEY_CURRENT_USER,_
 	    strKeyOutlookSetupPath,strFirstRun2ValueName

	'Save the applied ProfileVersion if larger than 1.
	If ProfileVersion > 1 Then
	    objRegistry.CreateKey HKEY_CURRENT_USER,strKeyProfileVersionPath
	    objRegistry.SetDWORDValue HKEY_CURRENT_USER,_
		strKeyProfileVersionPath,strProfileVersionName,ProfileVersion
	End If

    Else 
        Wscript.Echo "Crucial file in script could not be found." &vbNewLine & _
        "Please contact your system administrator." 
    End If

Else
    'The mail profile already exists so there is no need to launch Outlook with the profile switch.
    'Of course you are free to do something else here with the knowledge that the mail profile exists.
End If

'Cleaup
Set objRegistry = Nothing
Set objFSO = Nothing
' FileName:    114_cf.vbs
' Usage:       for login by all CF Joe Foss Field users to Area52 domain


'Variable Declaration/Definition
Dim objSysInfo		'Returns AD System Information
Dim objUser			'Sets user object variable
Dim strUserDN		'Holds user's Distinguished Name
Dim objGroupList	'Holds user's Groups
Dim strGroupName	'AD group name string
Dim strDriveLetter	'drive letter string
Dim strUNC			'UNC path to server share string
Dim objFSO			'file system object
Dim objNetwork		'network object
Dim objShell		'Creates Windows Shell Object

'Set Variables 
set objGroupList = CreateObject("Scripting.Dictionary")
set objNetwork = WScript.CreateObject("WScript.Network")
set objFSO = createobject("Scripting.FileSystemObject")
Set objShell = WScript.CreateObject("WScript.Shell")

'Collect User Account Info
Set objSysInfo = CreateObject("ADSystemInfo")
strUserDN = objSysInfo.UserName
strUserDN = Replace(strUserDN, "/", "\/")
Set objUser = GetObject("LDAP://" & strUserDN)

'Enumerate Group Membership for user
EnumerateGroups(objUser)
'--------------------------------------------------------------------------------

'Map Common Drives
'MapDrive "U:","\\servername\share"

'Specific group membership mappings
If MemberOf(objGroupList, "SecurityGroup") Then
	MapDrive "U:", "\\servername\share"
End If

'--------------------------------------------------------------------------------
'Enumerate Groups
Sub EnumerateGroups(ByVal objADObject)
	' Recursive subroutine to enumerate user's group membership to include nested groups
    Dim colstrGroups, objGroup, j
    objGroupList.CompareMode = vbTextCompare
    colstrGroups = objADObject.memberOf
    If (IsEmpty(colstrGroups) = True) Then
        Exit Sub
    End If
    If (TypeName(colstrGroups) = "String") Then
        colstrGroups = Replace(colstrGroups, "/", "\/")
        Set objGroup = GetObject("LDAP://" & colstrGroups)
        If (objGroupList.Exists(objGroup.sAMAccountName) = False) Then
            objGroupList.Add objGroup.sAMAccountName, True
            'Wscript.Echo objGroup.distinguishedName
			'Wscript.Echo objGroup.sAMAccountName
            Call EnumerateGroups(objGroup)
        End If
        Set objGroup = Nothing
        Exit Sub
    End If
    For j = 0 To UBound(colstrGroups)
        colstrGroups(j) = Replace(colstrGroups(j), "/", "\/")
        Set objGroup = GetObject("LDAP://" & colstrGroups(j))
        If (objGroupList.Exists(objGroup.sAMAccountName) = False) Then
            objGroupList.Add objGroup.sAMAccountName, True
            'Wscript.Echo objGroup.distinguishedName
			'Wscript.Echo objGroup.sAMAccountName
            Call EnumerateGroups(objGroup)
        End If
    Next
    set objGroup = Nothing
	set colstrGroups = Nothing
	set j = nothing
End Sub

'Checks for group membership against the list of groups
Function MemberOf (objGroupList, strGroupName)
	MemberOf = cbool(objGroupList.Exists(strGroupName))
End Function

'Maps Network Drive
Sub MapDrive(strDriveLetter, strUNC)
	'objNetwork.MapNetworkDrive strDriveLetter, strUNC
	If objFSO.DriveExists(strDriveLetter) Then
		objNetwork.RemoveNetworkDrive strDriveLetter, True, True
		objNetwork.MapNetworkDrive strDriveLetter, strUNC
	Else
		objNetwork.MapNetworkDrive strDriveLetter, strUNC, True
	End If
End Sub


'--------------------------------------------
' Custom SUBROUTINES & FUNCTIONS
'--------------------------------------------



'--------------------------------------------------------------------------------
'Clear memory
set objSysInfo = Nothing
set objUser = Nothing
set strUserDN = Nothing
set objGroupList = Nothing
set strGroupName = Nothing
set strDriveLetter = Nothing
set strUNC = Nothing
set objFSO = Nothing
set objNetwork = Nothing
set objShell = Nothing
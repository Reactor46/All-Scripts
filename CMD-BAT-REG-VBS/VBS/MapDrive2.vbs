' SOURCE: https://community.spiceworks.com/scripts/show/379-map-network-drive-if-user-memberof-group-vbscript
'start script

On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")
Set objNetwork = CreateObject("Wscript.Network")

'find user name

strUserPath = "LDAP://" & objSysInfo.UserName
Set objUser = GetObject(strUserPath)

'find user group's

For Each strGroup in objUser.MemberOf
strGroupPath = "LDAP://" & strGroup
Set objGroup = GetObject(strGroupPath)
strGroupName = objGroup.CN

' if user member of a group then map network drive

Select Case strGroupName
Case "GROUP NAME 1"
objNetwork.MapNetworkDrive "DRIVE LETTER:", "NETWORK PATH"

Case "GROUP NAME 2" 
objNetwork.MapNetworkDrive "DRIVE LETTER:", "NETWORK PATH"

Case "GROUP NAME 3"
objNetwork.MapNetworkDrive "DRIVE LETTER:", "NETWORK PATH"
'
End Select
Next

'end script

Found on Spiceworks: https://community.spiceworks.com/scripts/show/379-map-network-drive-if-user-memberof-group-vbscript?utm_source=copy_paste&utm_campaign=growth
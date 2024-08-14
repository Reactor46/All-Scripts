
Dim objWshFso, strTopFolder, colFolders, strFolder, objLogFile, strHomeShare 

Dim sSeparator 
sSeparator = ";" ' vbTab 

Sub createLog(strLogName) 
    'If (objWshFso.FileExists(strLogName)) = -1 Then 
    '    Set objLogFile = objWshFso.OpenTextFile(strLogName, 8) 
    'Else 
        Set objLogFile = objWshFso.CreateTextFile(strLogName, True) 
    'End If 
	
	sLogLine = "Name" & sSeparator & "ShortName" & sSeparator & "Type" & sSeparator & "IsReparsePoint" & sSeparator & "Extension" & sSeparator & "Path" & sSeparator & "ShortPath" & sSeparator & "Size" & sSeparator & "OwnerAccount" & sSeparator & "OwnerSID" & sSeparator & "IncludeInheritablePermissions" & sSeparator & "HasNotInheritedPermissions" & sSeparator & "SecurityDescriptorSDDL" & sSeparator & "SecurityDescriptorXml" & sSeparator & "CreationTime" & sSeparator & "LastAccessTime" & sSeparator & "LastWriteTime" & sSeparator & "ErrorFlag"
	objLogFile.WriteLine sLogLine

End Sub 


Sub TraverseFolders(fldr)

'Attributes Property | DateCreated Property | DateLastAccessed Property | DateLastModified Property 
'| Drive Property | Name Property | ParentFolder Property | Path Property | ShortName Property |
' ShortPath Property | Size Property | Type Property

	wscript.echo "Scanning " & fldr.path
	
	'Name	ShortName	Type	IsReparsePoint	Extension	Path	ShortPath	Size	OwnerAccount	OwnerSID	IncludeInheritablePermissions	HasNotInheritedPermissions	SecurityDescriptorSDDL	SecurityDescriptorXml	CreationTime	LastAccessTime	LastWriteTime	ErrorFlag
	'calc.exe	calc.exe	File	FALSE	.exe	\\SP2013\Branch1Share\calc.exe	\\SP2013\Branch1Share\calc.exe	918528	BUILTIN\Administrators	S-1-5-32-544	TRUE	FALSE	O:BAD:AI(A;ID;FA;;;SY)(A;ID;FA;;;BA)(A;ID;0x1200a9;;;BU)	<FileSystemItemAccessRules><FileSystemItemAccessRule AccessControlType="Allow" IdentityAccount="NT AUTHORITY\SYSTEM" IdentitySID="S-1-5-18" FileSystemRights="FullControl" IsInherited="True" InheritanceFlags="None" PropagationFlags="None" /><FileSystemItemAccessRule AccessControlType="Allow" IdentityAccount="BUILTIN\Administrators" IdentitySID="S-1-5-32-544" FileSystemRights="FullControl" IsInherited="True" InheritanceFlags="None" PropagationFlags="None" /><FileSystemItemAccessRule AccessControlType="Allow" IdentityAccount="BUILTIN\Users" IdentitySID="S-1-5-32-545" FileSystemRights="ReadAndExecute, Synchronize" IsInherited="True" InheritanceFlags="None" PropagationFlags="None" /></FileSystemItemAccessRules>	03/07/2013 17.12	03/07/2013 17.12	14/07/2009 03.38	FALSE

	For Each file in fldr.Files
		'Name
		sFileName = file.Name
		'ShortName
		sFileShortName = file.ShortName
		'Type
		sFileType = file.Type
		'IsReparsePoint
		sFileIsReparsePoint = "N/A"
		'Extension
		sExtension = objWshFso.GetExtensionName(sFileName)
		'Path
		sPath = file.Path
		'ShortPath
		sShortPath = file.ShortPath
		'Size
		sSize = file.Size
		'OwnerAccount
		sOwnerAccount = "N/A"
		'OwnerSID
		sOwnerSID = "N/A"
		'IncludeInheritablePermissions
		sIncludeInheritablePermissions = "N/A"
		'HasNotInheritedPermissions
		sHasNotInheritedPermissions = "N/A"
		'SecurityDescriptorSDDL
		sSecurityDescriptorSDDL = "N/A"
		'SecurityDescriptorXml
		sSecurityDescriptorXml = "N/A"
		'CreationTime
		sCreationTime = file.DateCreated
		'LastAccessTime
		sLastAccessTime = file.DateLastAccessed
		'LastWriteTime
		sLastWriteTime = file.DateLastModified
		'ErrorFlag
		sErrorFlag = "N/A"

		sLogLine = sFileName & sSeparator & sFileShortName  & sSeparator & sFileType & sSeparator & sFileIsReparsePoint & sSeparator & sExtension & sSeparator & sPath & sSeparator & sShortPath & sSeparator & sSize & sSeparator & sOwnerAccount & sSeparator & sOwnerSID & sSeparator & sIncludeInheritablePermissions & sSeparator & sHasNotInheritedPermissions & sSeparator & sSecurityDescriptorSDDL & sSeparator & sSecurityDescriptorXml & sSeparator & sCreationTime & sSeparator & sLastAccessTime & sSeparator & sLastWriteTime & sSeparator & sErrorFlag 
		wscript.echo sLogLine
		objLogFile.WriteLine sLogLine

	Next

	For Each sf In fldr.SubFolders
		TraverseFolders sf 
	Next

End Sub


sub Main()
 
	if (wscript.arguments.count > 1) then 
		strHomeShare = wscript.arguments(0)
		strLogName = wscript.arguments(1)
		wscript.echo "   Share: " & strHomeShare
		wscript.echo "Log name: " & strLogName 
	else
		wscript.echo "Uso: " & wscript.scriptname & " <share path> <log file name>"
		wscript.quit
	end if

	Set objWshFso = CreateObject("Scripting.FilesystemObject") 

	Call createLog(strLogName) 
	
	'On Error Resume Next 

	set oTopFolder = objWshFso.GetFolder(strHomeShare) 
	if (err.number <> 0) then
		wscript.echo "Error opening share: " & err.description
		wscript.quit
	end if

	Call TraverseFolders(oTopFolder)

	Set objWshFso = Nothing 

end Sub


Main
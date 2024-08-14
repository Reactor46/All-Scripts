'---------------------------------------------------------------------------------
' The sample scripts are not supported under any Microsoft standard support
' program or service. The sample scripts are provided AS IS without warranty
' of any kind. Microsoft further disclaims all implied warranties including,
' without limitation, any implied warranties of merchantability or of fitness for
' a particular purpose. The entire risk arising out of the use or performance of
' the sample scripts and documentation remains with you. In no event shall
' Microsoft, its authors, or anyone else involved in the creation, production, or
' delivery of the scripts be liable for any damages whatsoever (including,
' without limitation, damages for loss of business profits, business interruption,
' loss of business information, or other pecuniary loss) arising out of the use
' of or inability to use the sample scripts or documentation, even if Microsoft
' has been advised of the possibility of such damages.
'---------------------------------------------------------------------------------


USFpath  = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Favorites"
SFpath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Favorites"
Set objshell = CreateObject("WScript.Shell")
UserProfile = objshell.ExpandEnvironmentStrings("%USERPROFILE%")
OneDriveF = UserProfile & "\OneDrive\Favorites"
OneDrive = UserProfile & "\OneDrive\"
SkyDriveF = UserProfile & "\SkyDrive\Favorites"
SkyDrive = UserProfile & "\SkyDrive\"
Default = "%USERPROFILE%\Favorites"
DefaultF = UserProfile & "\Favorites"
USFvalue = objshell.RegRead(USFpath)
If USFvalue = OneDriveF Then 
	WScript.Echo "you have set favorites to OneDrive"
ElseIf USFvalue =  SkyDriveF Then 
	WScript.Echo "you have set favorites to OneDrive"
Else 

		Set FSO = CreateObject("Scripting.FileSystemObject")
		
		If FSO.FolderExists(SkyDrive) = True  Then 
			
			FSO.CopyFolder DefaultF, SkyDrive , True 
			objshell.RegWrite USFpath, SkyDriveF
			objshell.RegWrite SFpath,  SkyDriveF
			WScript.Echo "Move IE favorites to OneDrive successfully.You need to restart IE to take effect."
		ElseIf FSO.FolderExists(OneDrive) = True Then 
			
			FSO.CopyFolder DefaultF, OneDrive , True 
			objshell.RegWrite USFpath, OneDriveF
			objshell.RegWrite SFpath, OneDriveF 
			WScript.Echo "Move IE favorites to OneDrive successfully.You need to restart IE to take effect."
		Else 	 
			WScript.Echo "Not find OneDrive installed."
		End If 
End If  
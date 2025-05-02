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
Option Explicit 

Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE 
Dim strComputer, strKey, CurrentDirectory, Filepath
strComputer = "." 
strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" 
'Get script current folder
CurrentDirectory = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))
'set result file path 
Filepath = CurrentDirectory & "\InstalledSoftware.txt"
Dim FSO, TextFile, objReg, strSubkey, arrSubkeys 
'create filesystem object 
Set FSO = CreateObject("scripting.FileSystemObject")
'Create new text file 
Set TextFile = FSO.CreateTextFile(Filepath)
'Get WMI object 
Set objReg = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv") 
objReg.EnumKey HKLM, strKey, arrSubkeys 
Textfile.WriteLine  "Installed Applications: " 
Textfile.WriteLine
'Loop registry key.
Dim DisplayName,DisplayVersion, InstallDate, EstimatedSize, UninstallString
For Each strSubkey In arrSubkeys 
	objReg.GetStringValue HKLM, strKey & strSubkey, "DisplayName" , DisplayName
	If DisplayName <> "" Then 
		Textfile.WriteLine  "Display Name  : " & DisplayName
		objReg.GetStringValue HKLM, strKey & strSubkey, "DisplayVersion" , DisplayVersion
		Textfile.WriteLine  "Version       : " & DisplayVersion
		objReg.GetStringValue HKLM, strKey & strSubkey, "InstallDate", InstallDate
		Textfile.WriteLine  "InstallDate   : " & InstallDate 
		objReg.GetDWORDValue HKLM, strKey & strSubkey, "EstimatedSize" , EstimatedSize
		If  EstimatedSize <> "" Then 
			Textfile.WriteLine  "Estimated Size: " & Round(EstimatedSize/1024, 3) & " megabytes" 
		Else 
			Textfile.WriteLine  "Estimated Size: " 
		End If 
		objReg.GetStringValue HKLM, strKey & strSubkey, "UninstallString", UninstallString
		Textfile.WriteLine  "Uninstall     :" & UninstallString
		Textfile.Writeline
	End If 
Next 

TextFile.Close 
WScript.Echo "Generate 'InstalledSoftware.txt' successfully."
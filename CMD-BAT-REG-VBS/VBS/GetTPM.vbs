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
On Error Resume Next

Dim objFSO, strComputer,txtFile,TxtPath
Dim count ,objshell
Count = WScript.Arguments.Count
Select Case Count 
Case 1 
	Txtpath = WScript.Arguments(0)	
	Set objshell = CreateObject("wscript.shell")
	strComputer = objshell.ExpandEnvironmentStrings("%Computername%")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(TxtPath) Then 
		Set txtFile = objFSO.OpenTextFile(TxtPath , 8, True)
	Else
		Set txtFile = objFSO.CreateTextFile(TxtPath)
	End If 
	Dim  sConnection ,oTpmWmi, objTpm
	sConnection = "winmgmts:{impersonationLevel=impersonate,authenticationLevel=pktPrivacy}!root\cimv2\Security\MicrosoftTpm"
	Set oTpmWmi = GetObject(sConnection) 
	Set objTpm = oTpmWmi.Get("Win32_Tpm=@") 
	If err.number <> 0 then 
        txtFile.WriteLine "Failed to get TPM status from " & strComputer
	Else 
		txtFile.WriteLine  " "
		txtFile.writeline  "ManufacturerId             : " & strComputer
		txtFile.writeline  "ManufacturerId             : " & objTpm.ManufacturerId
		txtFile.writeline  "ManufacturerVersion        : " & objTpm.ManufacturerVersion
		txtFile.writeline  "ManufacturerVersionInfo    : " & objTpm.ManufacturerVersionInfo
		txtFile.writeline  "PhysicalPresenceVersionInfo: " & objTpm.PhysicalPresenceVersionInfo
		txtFile.writeline  "SpecVersion                : " & objTpm.SpecVersion
		txtFile.writeline  "IsActivated_InitialValue   : " & objTpm.IsActivated_InitialValue
		txtFile.writeline  "IsEnabled_InitialValue     : " & objTpm.IsEnabled_InitialValue
		txtFile.writeline  "IsOwned_InitialValue       : " & objTpm.IsOwned_InitialValue
		txtFile.WriteLine  " "
	End If  
Case Else
	wscript.echo "Invalid argument." 	
End Select 


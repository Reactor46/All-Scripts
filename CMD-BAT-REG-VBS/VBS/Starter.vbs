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

If IsFileExist("CCSPPT2007.exe") Then
	Set ws = CreateObject("WScript.Shell")
	' For Screen (150 ppi)
	ws.Exec "CCSPPT2007.exe -s 1"
	' For Print (220 ppi)
	'ws.Exec "CCSPPT2007.exe -s 0"
	' For E-mail (96 ppi)
	'ws.Exec "CCSPPT2007.exe -s 2"
Else
	WScript.Echo "Please put ""CCSPPT2007.exe"" in the current folder!"
End If

Function IsFileExist(filePath)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.fileExists(filePath) Then
		IsFileExist = True
	Else
		IsFileExist = False
	End If
End Function
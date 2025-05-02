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
Sub main()
On Error Resume Next 
Dim ArgCount
ArgCount = WScript.Arguments.Count
Select Case  ArgCount	
	Case 1 'Check the count of arguments
		Dim FSO,Path,File,Num_1,Num_2
		Set FSO = CreateObject("Scripting.FilesystemObject")
		Path = WScript.Arguments(0)
		If FSO.FileExists(Path) Then
			Set File = FSO.GetFile(path)
			If  (File.Attributes Mod 2) = 1 Then  'Check if the Read-Only is selected, and remove it.
				File.Attributes = File.Attributes-1 
				If Err.Number <> 0 Then 
					MsgBox "Error :" & Path &" "& Err.Description
				Else 
					MsgBox "Remove successfully."
				End If 
					
			Else 
				MsgBox "The Read-Only attribute of file is not selected"
			End If 
		Else 
			RemoveSubFolder Path,Num_1,Num_2 
			MsgBox Num_2 & " files successed" & ", " & Num_1 & " files Failed"
		End If 
	Case Else 
		MsgBox "Please drag a file or a folder."
End Select 
End Sub 

'This function is to remove the Read-Only of all files in a folder and its subfolder
Function RemoveSubFolder(FolderPath,Num_1,Num_2)
	On Error Resume Next 
	Dim FSObject,Folder
	Dim subFolder,File
	Num_1 = 0
	Num_2 = 0
	Set FSObject = CreateObject("Scripting.FilesystemObject")
	Set Folder = FSObject.GetFolder(FolderPath) 
	For Each  subFolder In Folder.SubFolders 'Loop the subfolder in the folder
		FolderPath = subFolder.Path 
		RemoveSubFolder FolderPath,Num_1,Num_2
	Next 
	For Each  File In Folder.Files 'Remove the Read-Only attribute of files in the folder
		If  (File.Attributes Mod 2) = 1 Then 
			File.Attributes = File.Attributes-1 
			If Err.Number <> 0 Then 
				MsgBox  "Error :" & File.Path &" "& Err.Description
				Num_1 = Num_1 + 1
			Else 
				Num_2 = Num_2 + 1 
			End If 
		End If 
		Err.Clear 
	Next 
	Set FSObject = Nothing 
End Function 

Call main 
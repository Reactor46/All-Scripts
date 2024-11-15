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
'################################################
'This script is to export Visio files to PNG files
'################################################

Sub main()
Dim ArgCount
ArgCount = WScript.Arguments.Count
Select Case ArgCount 
	Case 1	
		MsgBox "Please please make sure you drag a Visio file or folder that contains some Visio files,and press 'OK' to continue",,"Information"
		Dim VisioPaths,objshell
		VisioPaths = WScript.Arguments(0)

		Set objshell = CreateObject("scripting.filesystemobject")
		If objshell.FolderExists(VisioPaths) Then  'Check if the object is a folder
			Dim flag,FileNumber
			flag = 0 
			FileNumber = 0 	
			Dim Folder,VisioFiles,VisioFile		
			Set Folder = objshell.GetFolder(VisioPaths)
			Set VisioFiles = Folder.Files
			For Each VisioFile In VisioFiles  'loop the files in the folder
				FileNumber = FileNumber + 1 
				VisioFile = VisioFile.Path
				If GetVisioFile(VisioFile) Then  'if the file is Visio file, then convert it 
					ConvertVisioToPNG VisioFile
					flag = flag + 1
				End If 	
			Next 
			WScript.Echo "Totally " & FileNumber & " files in the folder and convert " & flag & " Visio file(s) to PNG fles."	
		Else 
			If GetVisioFile(VisioPaths) Then  'if the object is a file,then check if the file is a Visio file and convert it 
				Dim VisioPath
				VisioPath = VisioPaths
				ConvertVisioToPNG VisioPath
			Else 
				WScript.Echo "Please please make sure you drag a Visio file or a folder that contains some Visio files."
			End If  
		End If 
			
	Case  Else 
	 	WScript.Echo "Please please make sure you drag a Visio file or a folder that contains some Visio files."
End Select 
End Sub 

Function ConvertVisioToPNG(VisioFile)  'This function is to convert a Visio file to PNG file
	Dim objshell,ParentFolder,BaseName,Visioapp,Visio

	Set Visioapp = CreateObject("Visio.Application")
	Visioapp.Visible = False
	Set Visio = Visioapp.Documents.Open(VisioFile)
	Set Pages = Visioapp.ActiveDocument.Pages
	
	Set objshell= CreateObject("scripting.filesystemobject")
	ParentFolder = objshell.GetParentFolderName(VisioFile) 'Get the current folder path
	BaseName = objshell.GetBaseName(VisioFile) 'Get the file name
	
	Dim PageName,Page,Pages
	For Each Page In Pages
		PageName = Page.Name
		Page.Export(parentFolder & "\" & BaseName & "-" & PageName & ".png")
	Next
	
	Visio.Close
	Visioapp.Quit
	Set objshell = Nothing 
End Function 

Function GetVisioFile(VisioFile) 'This function is to check if the file is a Visio file
	Dim objshell
	Set objshell= CreateObject("scripting.filesystemobject")
	Dim Arrs ,Arr
	Arrs = Array("vsdx","vssx","vstx","vxdm","vssm","vstm","vsd","vdw","vss","vst")
	
	Dim blnIsVisioFile,FileExtension
	blnIsVisioFile = False 
	FileExtension = objshell.GetExtensionName(VisioFile)  'Get the file extension
	For Each Arr In Arrs
		If InStr(UCase(FileExtension),UCase(Arr)) <> 0 Then 
			blnIsVisioFile= True
			Exit For 
		End If 
	Next 
	GetVisioFile = blnIsVisioFile
	Set objshell = Nothing 
End Function 

Call main 
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
Dim ArgCount
ArgCount = WScript.Arguments.Count
Select Case ArgCount 
	Case 1	
		MsgBox "Please ensure Word documents are saved,if that press 'OK' to continue",,"Warning"
		Dim TxtFilePaths,objshell
		TxtFilePaths = WScript.Arguments(0)
		StopWordApp
		Set objshell = CreateObject("scripting.filesystemobject")
		If objshell.FolderExists(TxtFilePaths) Then  'Check if the object is a folder
			Dim flag,FileNumber
			flag = 0 
			FileNumber = 0 	
			Dim Folder,DocFiles,DocFile		
			Set Folder = objshell.GetFolder(TxtFilePaths)
			Set DocFiles = Folder.Files
			For Each DocFile In DocFiles  'loop the files in the folder
				FileNumber=FileNumber+1 
				TxtFilePath = DocFile.Path
				If GetWordFile(TxtFilePath) Then  'if the file is text file, then convert it 
					ConvertTxtToWord TxtFilePath
					flag=flag+1
				End If 	
			Next 
			WScript.Echo "Totally " & FileNumber & " files in the folder and convert " & flag & " '.txt' files to Word documents."
				
		Else 
			If GetWordFile(TxtFilePaths) Then  'if the object is a file,then check if the file is a text file.if that, convert it 
				Dim TxtFilePath
				TxtFilePath = TxtFilePaths
				ConvertTxtToWord TxtFilePath
			Else 
				WScript.Echo "Please drag a '.txt' file or a folder with '.txt' files."
			End If  
		End If 
			
	Case  Else 
	 	WScript.Echo "Please drag a '.txt' file or a folder with '.txt' file."
End Select 
End Sub 

Function ConvertTxtToWord(TxtFilePath)
Dim objshell,ParentFolder,BaseName,wordapp,doc,WordFilePath,objDoc
Dim objSelection
Set objshell= CreateObject("scripting.filesystemobject")
ParentFolder = objshell.GetParentFolderName(TxtFilePath) 'Get the current folder path
BaseName = objshell.GetBaseName(TxtFilePath) 'Get the text file name
WordFilePath = parentFolder & "\" & BaseName & ".docx"  
Dim objWord
set objWord = CreateObject("Word.Application")
Set objDoc=objWord.Documents.Add()
With objWord
   .Visible = False 
End With
set objSelection=objWord.Selection
objSelection.InsertFile TxtFilePath 
objDoc.saveas(WordFilePath)
objDoc.close
objWord.Quit
End Function 

Function GetWordFile(TxtFilePath) 'This function is to check if the file is a '.txt' file
	Dim objshell
	Set objshell= CreateObject("scripting.filesystemobject")
	Dim Arrs ,Arr
	Arrs = Array("txt")
	Dim blnIsDocFile,FileExtension
	blnIsDocFile= False 
	FileExtension = objshell.GetExtensionName(TxtFilePath)  'Get the file extension
	For Each Arr In Arrs
		If InStr(UCase(FileExtension),UCase(Arr)) <> 0 Then 
			blnIsDocFile= True
			Exit For 
		End If 
	Next 
	GetWordFile = blnIsDocFile
	Set objshell = Nothing 
End Function 


Function StopWordApp 'This function is to stop the Word application
	Dim strComputer,objWMIService,colProcessList,objProcess 
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	'Get the WinWord.exe
	Set colProcessList = objWMIService.ExecQuery _
		("SELECT * FROM Win32_Process WHERE Name = 'Winword.exe'")
	For Each objProcess in colProcessList
		'Stop it
		objProcess.Terminate()
	Next
End Function 

Call main 


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
'This script is to convert Word documents to PDF files
'################################################
Sub main()
Dim ArgCount
ArgCount = WScript.Arguments.Count
Select Case ArgCount 
	Case 1	
		MsgBox "Please ensure Word documents are saved,if that press 'OK' to continue",,"Warning"
		Dim DocPaths,objshell
		DocPaths = WScript.Arguments(0)
		StopWordApp
		Set objshell = CreateObject("scripting.filesystemobject")
		If objshell.FolderExists(DocPaths) Then  'Check if the object is a folder
			Dim flag,FileNumber
			flag = 0 
			FileNumber = 0 	
			Dim Folder,DocFiles,DocFile		
			Set Folder = objshell.GetFolder(DocPaths)
			Set DocFiles = Folder.Files
			For Each DocFile In DocFiles  'loop the files in the folder
				FileNumber=FileNumber+1 
				DocPath = DocFile.Path
				If GetWordFile(DocPath) Then  'if the file is Word document, then convert it 
					ConvertWordToPDF DocPath
					flag=flag+1
				End If 	
			Next 
			WScript.Echo "Totally " & FileNumber & " files in the folder and convert " & flag & " Word Documents to PDF fles."
				
		Else 
			If GetWordFile(DocPaths) Then  'if the object is a file,then check if the file is a Word document.if that, convert it 
				Dim DocPath
				DocPath = DocPaths
				ConvertWordToPDF DocPath
			Else 
				WScript.Echo "Please drag a word document or a folder with word documents."
			End If  
		End If 
			
	Case  Else 
	 	WScript.Echo "Please drag a word document or a folder with word documents."
End Select 
End Sub 

Function ConvertWordToPDF(DocPath)  'This function is to convert a word document to pdf file
	Dim objshell,ParentFolder,BaseName,wordapp,doc,PDFPath
	Set objshell= CreateObject("scripting.filesystemobject")
	ParentFolder = objshell.GetParentFolderName(DocPath) 'Get the current folder path
	BaseName = objshell.GetBaseName(DocPath) 'Get the document name
	PDFPath = parentFolder & "\" & BaseName & ".pdf" 
	Set wordapp = CreateObject("Word.application")
	Set doc = wordapp.documents.open(DocPath)
	doc.saveas PDFPath,17
	doc.close
	wordapp.quit
	Set objshell = Nothing 
End Function 

Function GetWordFile(DocPath) 'This function is to check if the file is a Word document
	Dim objshell
	Set objshell= CreateObject("scripting.filesystemobject")
	Dim Arrs ,Arr
	Arrs = Array("doc","docx")
	Dim blnIsDocFile,FileExtension
	blnIsDocFile= False 
	FileExtension = objshell.GetExtensionName(DocPath)  'Get the file extension
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
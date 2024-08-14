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

'This script shows how to insert images in a folder into a word document

Option Explicit 
Sub Main()
On Error Resume Next 
Err.clear 
Dim ArgCount
ArgCount = WScript.Arguments.Count 'Get the number of the arguments
Select Case ArgCount
	Case 2
		Const END_OF_STORY = 6
		Const MOVE_SELECTION = 0
		Dim objWord,DocPath 
		Dim objFSO,Folder,FolderPath,objDoc,ObjSelection
		Dim Image,imagepath
		Err = 0
		DocPath =WScript.Arguments(0)  'Get the path of the word document
		'Check if the first file extension is word document
		If CheckDocFileExtension(DocPath) Then 
			If Err.number = 0 Then 
				Set objWord = CreateObject("Word.Application") 
				objWord.Visible = False 
				'Get the path of the specified folder with images
				FolderPath =WScript.Arguments(1)
				Set objFSO = CreateObject("Scripting.Filesystemobject")
				Set Folder = objFSO.GetFolder(FolderPath)
				'verify if the second object is a foler.If not, quit
				If Err.number =0  Then 
					Set objDoc = objWord.Documents.open(DocPath) 'Open the word document
					Set objSelection = objWord.Selection
					For Each Image In Folder.Files
						imagepath = image.Path 
						If CheckiImageExtension(ImagePath) = True Then 
							'Insert the images into the word document
							objSelection.EndKey END_OF_STORY,MOVE_SELECTION
							objSelection.InlineShapes.AddPicture(imagepath)
							objSelection.insertbreak  'Insert a pagebreak 
						End If 	
					Next 
					'Delete the last blank page
					objSelection.TypeBackspace
					objSelection.TypeBackspace
					objDoc.save
					objWord.quit
					WScript.Echo "Insert images successfully!"

					Set objWord = Nothing 
					Set objFSO = Nothing 
					Set Folder = Nothing 
					Set objDoc = Nothing 
					Set objSelection = Nothing 
				Else
					WScript.Echo "Please Drag one Word Document and one folder with images"
					Err.clear
					WScript.Quit
				End If 
			Else 
				WScript.echo "The file input is not a Word document"
			End If 
		Else 
			WScript.Echo "The file input is not a Word document"
		End If 
	Case Else 
		WScript.Echo "Please Drag one Word Document and one folder with images"
End Select 

End Sub 

' ##################################################################
' Check if the given file extension is Word document.
' ##################################################################
Function CheckDocFileExtension(wdPath)
	Dim varArray		' An array contains word document file extensions.
	Dim varEach			' Each word document file extension.
	Dim blnIsPptFile	' Whether the file extension is word document file extension.
	Dim objFSO,file,FileExtension
	Set objFSO = CreateObject("Scripting.Filesystemobject")
	Set file = objFSO.GetFile(wdPath) 
	FileExtension = file.name
	blnIsPptFile = False	
	If FileExtension <> "" Then 
		varArray = Array(".doc",".docx")
		For Each varEach In varArray
			If InStrRev(FileExtension,varEach) <> 0 Then
				blnIsPptFile = True
				Exit For
			End If
		Next
	End If
	CheckDocFileExtension = blnIsPptFile
	Set objFSO = Nothing 
	Set file = Nothing 
End Function

' ##################################################################
' Check if the file in folder is image file.
' ##################################################################
Function CheckiImageExtension(ImagePath)
	Dim varArray		' An array contains iamge file extensions.
	Dim varEach			' Each iamge file extension.
	Dim blnIsPptFile	' Whether the file extension is image file extension.
	Dim objFSO,file,FileExtension
	Set objFSO = CreateObject("Scripting.Filesystemobject")
	Set File = objFSO.GetFile(ImagePath)
	FileExtension = File.name
	blnIsPptFile = False
	If FileExtension <> "" Then
		varArray = Array(".emf", ".wmf",".jpg",".jpeg",".jfif",".png",".jpe",".bmp",".dib",".rle",_
		                 ".gif",".emz",".wmz",".pcz",".tif",".tiff",".eps",".pct",".pict",".wpg")
		For Each varEach In varArray
			If InStrRev(UCase(FileExtension),UCase(varEach)) <> 0 Then
				blnIsPptFile = True
				Exit For
			End If
		Next
	End If
	CheckiImageExtension = blnIsPptFile
	Set objFSO = Nothing 
	Set file = Nothing 
End Function

Call Main 
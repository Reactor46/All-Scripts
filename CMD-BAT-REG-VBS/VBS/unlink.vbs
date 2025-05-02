
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

Set oWord = CreateObject("Word.Application")
oWord.Visible = False
oWord.Documents.Open WScript.Arguments(0)
For Each oSection In oWord.ActiveDocument.Sections
	For Each oFooter In oSection.Footers
		oFooter.Range.Fields.Unlink
	Next 
Next 
oWord.ActiveDocument.Fields.Unlink
oWord.Documents.Save
oWord.Documents.Close
oWord.Quit
ConvertWordToPDF WScript.Arguments(0)
Set oWord = Nothing
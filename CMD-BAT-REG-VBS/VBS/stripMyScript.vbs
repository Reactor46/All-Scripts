'=*=*=*=*=*=*=*=*=*=*=*=*=
' StripMyScript.vbs
' Created by Assaf Miron
' Date : 04/12/06
' Description : This Script Replaces Words that are not Allowed to go outside of the Organization
' You can Define a List of Words and there Replacemenst and Run This Script on Any File
' in order to strip it out of the elegal content
'=*=*=*=*=*=*=*=*=*=*=*=*=

Const ForReading = 1
Const ForWriting = 2

Dim objFSO
Dim arrEText() 'Array of Elegal Text
Dim arrAText() 'Array of Approved Text, Replaces the Elegal Text

Set objFSO = CreateObject("Scripting.FileSystemObject")

Sub Usage()
WScript.Echo "StripMyScript.vbs" & vbNewLine _
& vbTab & "by : Assaf Miron" & vbNewLine _
& "---------------------------------------------------------------------" & vbNewLine _
& " With This Script you can Strip almost any file of words that are not" & vbNewLine _
& " allowed to be exposed." & vbNewLine _
& " If you need to export a file from your company network but you don't" & vbNewLine _
& "  want to export any confidetial content, you can specify all the " & vbNewLine _
& "  Elegal words that you think may be in the File and Replace them with" & vbNewLine _
& "  Another Word." & vbNewLine _
& " All you need is to Create a TXT File in a format of :" & vbNewLine _
& vbTab & " SecretWord,ReaplcedWord" & vbNewLine _
& " and save it on a Server or on your computer." & vbNewLine _
& "---------------------------------------------------------------------" & vbNewLine & vbNewLine _
& " Syntax :" & vbNewLine _
& " StripMyScript.vbs /EX:C:\<MySecretWords.txt> /File:C:\<MyExportedFile>" & vbNewLine & vbNewLine _
& " <MySecretWords.txt> is the premade file with all the Words need to be Replaced" & vbNewLine _
& " <MyExportedFile> is a file you want to Strip out of Content." & vbNewLine _
& " Script can receive these Extenstions : txt,vbs,bat,log,csv,xls,xlsx,doc,docx,rtf" & vbNewLine & vbNewLine _
& " Remarks : The Script Should be run in CScript."


WScript.Quit
End Sub

Sub FileType(FileName)
' This sub recieves the File name and by the Extention opens it with the right program
Dim FileNameExtention
Dim objFile
Set objFile = objFSO.GetFile(FileName)
FileNameExtention = objFSO.GetExtensionName(objFile)
Select Case FileNameExtention
	Case LCase("txt")
		StripTextFile FileName
	Case LCase("vbs")
    StripTextFile FileName
	Case LCase("csv") 
		StripTextFile FileName
	Case LCase("log")
		StripTextFile FileName
	Case LCase("bat")
		StripTextFile FileName
	Case LCase("xls")
		StripExcelFile FileName
	Case LCase("xlsx")
		StripExcelFile FileName
	Case LCase("doc")
		StripWordFile FileName
  Case LCase("docx")
    StripWordFile FileName
	Case LCase("rtf")
		StripWordFile FileName
End Select
End Sub

Sub StripWordFile(strFilePath)
Const wdReplaceAll = 2
Dim I
Dim objWord,objDoc
Dim objSelection

Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Open(strFilePath)
Set objSelection = objWord.Selection


For I = 0 To UBound(arrEText)-1
	objSelection.Find.Text = arrEText(I)
	objSelection.Find.Forward = TRUE
    objSelection.Find.MatchWholeWord = True
	objSelection.Find.Replacement.Text = arrAText(I)
	objSelection.Find.Execute ,,,,,,,,,,wdReplaceAll
Next

objDoc.Save
objDoc.Close
objWord.Quit
Set objSelection = Nothing
Set objDoc = Nothing
Set objWord = Nothing
End Sub

Sub StripExcelFile(strFilePath)

Const xlCSV = 6

Dim objExcel,objWorkBook,objWorksheet
Dim intWorkSheet
Dim WS

Set objExcel = CreateObject("Excel.Application")
Set objWorkBook = objExcel.Workbooks.Open(strFilePath)
objExcel.DisplayAlerts = False

intWorkSheet = objWorkBook.Worksheets.Count

For WS = 1 To intWorkSheet
WScript.Echo ws
	Set objWorksheet = objWorkbook.Worksheets(WS)
	objWorksheet.SaveAs "c:\tempWorkSheet.csv", xlCSV
	StripTextFile "c:\tempWorkSheet.csv"
	objWorksheet.Delete
	Set objWorksheet = Nothing
	Set objWorksheet = objWorkBook.Worksheets.Add
	objWorksheet.Open "c:\tempWorkSheet.csv"
	objExcel.Save
	Set objWorksheet = Nothing
	objFSO.DeleteFile "c:\tempWorkSheet.csv",True
Next

objExcel.Quit
Set objWorkBook = Nothing
Set objExcel = Nothing

End Sub

Sub StripTextFile(strFilePath)
Dim objFile
Dim tmpString

Set objFile = objFSo.OpenTextFile(strFilePath,ForReading)
tmpString = objFile.ReadAll
objFile.Close
Set objFile = Nothing
tmpString = StripStringText(tmpString)
Set objFile = objFSo.OpenTextFile(strFilePath,ForWriting)
objFile.WriteLine tmpString
objFile.Close
Set objFile = Nothing

End Sub

Function StripStringText(strText)
Dim I
For I = 0 To UBound(arrEText)
WScript.Echo arrEText(I),arrAText(I)
	strText = Replace(LCase(strText),LCase(arrEText(i)),arrAText(I))
Next
StripStringText = strText
End Function


' ---- Main Code ----

Dim objTextFile
Dim ETextFile
Dim FilePath
Dim strLine
Dim i

If WScript.Arguments.Count < 2 Then _
	Usage()
If WScript.Arguments.Named("?") Then _
	Usage()

If (WScript.Arguments.Named.Exists("EX")) And (WScript.Arguments.Named.Exists("File")) Then
	With WScript.Arguments
		ETextFile = .Named("EX")
		FilePath = .Named("File")
	End With
End If

Set objTextFile = objFSO.OpenTextFile(ETextFile,ForReading)
i = 0
Do Until objTextFile.AtEndOfStream
	strLine = objTextFile.ReadLine
	arrTempString = Split(strLine,",")
	Redim Preserve arrEText(i)
	Redim Preserve arrAText(i)
	arrEText(i) = arrTempString(0)
	arrAText(i) = arrTempString(1)
	i = i + 1
Loop

objTextFile.Close
Set objTextFile = Nothing

FileType FilePath

Set objFSo = Nothing
Wscript.Echo "Done !"

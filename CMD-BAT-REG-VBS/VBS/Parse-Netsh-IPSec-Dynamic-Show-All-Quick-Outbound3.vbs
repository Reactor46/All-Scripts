'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
' Created by Assaf Miron
' Http://assaf.miron.googlepages.com
' Date : 02/04/2009
' Parse-Netsh-IPSec-Dynamic-Show-All-Quick-Outbound2.vbs
' Description : Parses the NETSH Output into Excel (Version 3)
' 				Gets the Output of the Command: 
'				NetSH Dynamic show All
'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Dim objFSO,objFile
Dim arrLines
Dim strLine
Dim objExcel,objWorkbook
Dim FileLoc
Dim intRow
Dim objDictionary

FileLoc = "C:\IPSecWeights.xls"

Sub ExcelHeaders()
	Set objRange = objExcel.Range("A1","G1")
	objRange.Font.Size = 12
	objRange.Interior.ColorIndex=15
	
	objexcel.cells(1,1)="Filter Name"
	objexcel.cells(1,2)="Source"
	objexcel.cells(1,3)="Destination"
	objexcel.cells(1,4)="Source Port"
	objexcel.cells(1,5)="Destination Port"
	objexcel.cells(1,6)="Protocol"
	objexcel.cells(1,7)="Direction"
End Sub

Function RegExFind(strText,strPattern)
	Dim regEx
	Dim match, Matches
	Dim arrMatches
	Dim i : i = 0
	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Global = True
	regEx.Pattern = strPattern
	
	Set matches = regEx.Execute(strText)
	ReDim arrMatches(Matches.Count)
	For Each match In Matches
		For Each SubMatch In match.Submatches
			arrMatches(i) = Submatch
			i = i + 1
		Next
	Next
	RegExFind = arrMatches
End Function


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(WScript.Arguments(0),ForReading)

Set objExcel = CreateObject("excel.application")
Set objWorkbook = objExcel.Workbooks.Open(FileLoc)

objExcel.Visible = True

ExcelHeaders ' Create Excel Headers

rePolicy = "Policy Name\s+:\s(.+)"
reSRCAddr = "Source Address\s+:\s(.+)"
reDSTAddr = "Destination Address\s+:\s(.+)"
reProtocol = "Protocol\s+:\s(.+)"
reSRCPort = "Source Port\s+:\s(.+)"
reDSTPort = "Destination Port\s+:\s(.+)"
reDirection = "Direction\s+:\s(.+)"

strText = objFile.ReadAll
objFile.Close

Dim arrPolicy, arrSRCAddr, arrDSTAddr, arrProtocol, arrSRCPort, arrDSTPort, arrDirection

arrPolicy = RegExFind(strText, rePolicy)
arrSRCAddr = RegExFind(strText, reSRCAddr)
arrDSTAddr = RegExFind(strText, reDSTAddr)
arrProtocol = RegExFind(strText, reProtocol)
arrSRCPort = RegExFind(strText, reSRCPort)
arrDSTPort = RegExFind(strText, reDSTPort)
arrDirection = RegExFind(strText, reDirection)

intRow = 2

For i = 0 To UBound(arrPolicy)
	objExcel.Cells(introw,1) = arrPolicy(i)
	objExcel.Cells(introw,2) = arrSRCAddr(i)
	objExcel.Cells(introw,3) = arrDSTAddr(i)
	objExcel.Cells(introw,4) = arrSRCPort(i)
	objExcel.Cells(introw,5) = arrDSTPort(i)
	objExcel.Cells(introw,6) = arrProtocol(i)
	objExcel.Cells(introw,7) = arrDirection(i)

	intRow = intRow + 1
Next

objFile.Close
objWorkbook.save
'objExcel.Quit
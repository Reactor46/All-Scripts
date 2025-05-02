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

'This script shows how to change the default paper size in Word 

Option Explicit
Sub main()
	On Error Resume Next	
	Dim Context
	Dim PaperSize
	Dim envUSER,objshell
	Set objshell = CreateObject("wscript.shell")
	'Get the current username
	envUSER = objshell.expandEnvironmentStrings("%username%")
	'Dim a variable to store the papersize number
	Context = "10x14 = 0 "_
			& vbNewline & "11x17 = 1"_
			& vbNewline & "Letter = 2"_
			& vbNewline & "LetterSmall = 3"_
			& vbNewline & "Legal = 4"_
			& vbNewline & "Executive = 5"_
			& vbNewline & "A3 = 6"_
			& vbNewline & "A4 = 7"_
			& vbNewline & "A4Small = 8"_			
			& vbNewline & "A5 = 9"_					
			& vbNewline & "B4 = 10"_					
			& vbNewline & "B5 = 11"_						
			& vbNewline & "CSheet = 12"_					
			& vbNewline & "DSheet = 13"_					
			& vbNewline & "ESheet = 14"_					
			& vbNewline & "FanfoldLegalGerman = 15"_		
			& vbNewline & "FanfoldStdGerman = 16"_			
			& vbNewline & "FanfoldUS = 17"_				
			& vbNewline & "Folio = 18"_					
			& vbNewline & "Ledger = 19"_					
			& vbNewline & "Note = 20"_						
			& vbNewline & "Quarto = 21"_					
			& vbNewline & "Statement = 22"_				
			& vbNewline & "Tabloid = 23"_					
			& vbNewline & "Envelope9 = 24"_				
			& vbNewline & "Envelope10 = 25"_				
			& vbNewline & "Envelope11 = 26"_				
			& vbNewline & "Envelope12 = 27"_				
			& vbNewline & "Envelope14 = 28"_				
			& vbNewline & "EnvelopeB4 = 29"_				
			& vbNewline & "EnvelopeB5 = 30"_				
			& vbNewline & "EnvelopeB6 = 31"_				
			& vbNewline & "EnvelopeC3 = 32"_				
			& vbNewline & "EnvelopeC4 = 33"_				
			& vbNewline & "EnvelopeC5 = 34"_				
			& vbNewline & "EnvelopeC6 = 35"_				
			& vbNewline & "EnvelopeC65 = 36"_				
			& vbNewline & "EnvelopeDL = 37"_				
			& vbNewline & "EnvelopeItaly = 38"_			
			& vbNewline & "EnvelopeMonarch = 39	"_		
			& vbNewline & "EnvelopePersonal = 40"_
			& vbNewline & "---------------------"_
			& vbNewline & "Please enter a number to set default paper size:"
	'Prompt the paper size number and get a number 
	PaperSize = InputBox(Context,"Paper Size Configuration")
	'Check if the value input is between 0 and 41
	If  IsNumeric(PaperSize)   Then
		If CLng(PaperSize) Then
			If PaperSize >= 0 And PaperSize <41 Then 
				'Stop the Word application
				Call StopWordApp
				Dim wdApp,Doc
				'Modify the word template document
				Set wdApp = Createobject("Word.Application")
				Set Doc = wdApp.documents.open("C:\Users\" & envUSER & "\AppData\Roaming\Microsoft\Templates\Normal.dotm")
				'Set the specified number as the default paper size 
				With wdApp.Selection.PageSetup
					.PaperSize = PaperSize
					.SetAsTemplateDefault
				End With
				Doc.Save
				Doc.Close 
				If Err.number <> 0 Then 
					wdApp.Quit
					msgbox Err.Description,,"Configuration Warning"
				Else 
					wscript.echo "Configure the default paper size successfully."
				End If 
				wdApp.Quit
				Call StopWordApp
				If Not IsNull(wdApp) Then Set wdApp = Nothing
			Else 
				msgbox "Invalid value ,please try again.",,"Warning"
			End If 
		Else 
			msgbox "Invalid value ,please try again.",,"Warning"
		End If 
	Else 
		msgbox "Invalid value ,please try again.",,"Warning"
	End If 
End Sub 

'This function is to stop the Word application
Function StopWordApp
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
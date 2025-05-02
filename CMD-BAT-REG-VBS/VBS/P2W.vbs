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

Call Main

' ################################################
' The starting point of execution for this script.
' ################################################
Sub Main()
	Const P2W_MSG_NOPARAM = "Please drag a PowerPoint presentation file onto this VBScript file."
	Const P2W_MSG_TASKEND = "Conversion complete."
	Const P2W_MSG_TASKBEGIN = "Starting the conversion..."
	Const P2W_MSG_DRAGMULTIPLE = "Too many files or folders."
	
	Dim wsArgNum
	
	' Get the number of parameters from command-line.
	wsArgNum = WScript.Arguments.Count
	
	Select Case wsArgNum
		Case 0
			WScript.Echo P2W_MSG_NOPARAM
		Case 1
			Dim strFullPath
			Dim strFileExtension
			
			strFullPath = WScript.Arguments(0)
			strFileExtension = GetFileExtensionFromFullPath(strFullPath)
			
			' /* If the file extension in the full path is a valid PPT file extension. */
			If CheckPptFileExtension(strFileExtension) = True Then
				Dim objWsh
				
				Set objWsh = CreateObject("WScript.Shell")
				objWsh.Popup P2W_MSG_TASKBEGIN, 1, "Message", 64
				
				' Begin to convert the PowerPoint presentation.
				Call SendPowerPoint2Word(strFullPath)
				
				objWsh.Popup P2W_MSG_TASKEND, , "Message", 4096 + 64
				Set objWsh = Nothing
			Else
				WScript.Echo P2W_MSG_NOPARAM
				WScript.Quit
			End If
		Case Else
			WScript.Echo P2W_MSG_DRAGMULTIPLE
			WScript.Quit
	End Select
End Sub

' #####################################
' Send PowerPoint presentation to Word.
' #####################################
Function SendPowerPoint2Word(FullPath)
	Dim wdApp				' Word application.
	Dim wdDoc				' Word document.
	Dim sldEach				' Each slide.
	Dim sldAll				' All slides.
	Dim spNotes				' Notes shapes.
	Dim spNotesPage			' All shapes in notes page.
	Dim pptPresentation		' PowerPoint presentation.
	Dim sldHeight			' Slide height.
	Dim sldWidth			' Slide width.
	Dim strFilePath			' File path.
	Dim strFileName			' File name with no extension.
	Dim strNotesText		' Notes text.
	Dim intPageNumber		' Page number in Word.
	
	intPageNumber = 0
	
	' /* Constants declaration. */
	Const wdPaperCustom = 41
	Const wdStory = 6
	Const wdCharacter = 1
	Const wdExtend = 1
	Const wdGoToPage = 1
	Const wdGoToNext = 2
	Const ppPlaceholderBody = 2
	
	' Get the file path from the given path.
	strFilePath = GetFilePathFromFullPath(FullPath)
	' Get the file name from the given path.
	strFileName = GetFileNameFromFullPath(FullPath)
	
	On Error Resume Next
	
	Set pptPresentation = GetObject(FullPath)
	
	' /* Get the slide size. */
	With pptPresentation.PageSetup
		sldHeight = .SlideHeight
		sldWidth = .SlideWidth
	End With
	
	Set wdApp = CreateObject("Word.Application")
	Set wdDoc = wdApp.Documents.Add
	
	' /* Page setup in Word. */
	With wdApp.Selection.PageSetup
		.LeftMargin = 0
		.RightMargin = 0
		.TopMargin = 0
		.BottomMargin = 0
		.PaperSize = wdPaperCustom
		.PageWidth = sldWidth
		.PageHeight = sldHeight
	End With
	
	' Reference to all slides in PowerPoint presentation.
	Set sldAll = pptPresentation.Slides
	
	With wdApp.Selection
		' /* Go through each slide object. */
		For Each sldEach in sldAll
			
			Set spNotesPage = sldEach.NotesPage.Shapes
			
			' /* Read notes in the current slide. */
			For Each spNotes In spNotesPage
				If spNotes.HasTextFrame Then
					If spNotes.PlaceholderFormat.Type = ppPlaceholderBody Then
						strNotesText = spNotes.TextFrame.TextRange.Text
						Exit For
					End If
				End If
			Next
			
			sldEach.Shapes.Range.Copy
			.Paste
			
			' To count the page number.
			intPageNumber = intPageNumber + 1
			
			.ShapeRange.Group
			.ShapeRange.Left = 0
			.ShapeRange.Ungroup
			
			' /* If the current slide has notes. */
			If strNotesText <> "" Then
				' Goto the first line of the current page in Word.
				.GoTo wdGoToPage, wdGoToNext, , intPageNumber
				' Set the selection text to space.
				.Text = Space(1)
				' Move the cursor to the right of the space.
				.MoveRight wdCharacter, 1
				' Add comments from the current slide's notes into the Word document.
				wdDoc.Comments.Add .Range, strNotesText
			End If
			
			.EndKey wdStory
			.InsertNewPage
		Next
		
		pptPresentation.Close
		pptPresentation.Application.Quit
		
		' /* To delete the last blank page in Word. */
		.TypeBackspace
		.TypeBackspace
		
		' /* Copy the newline char to reduce a large data on clipboard. */
		.MoveRight wdCharacter, 1, wdExtend
		.Copy
	End With
	
	wdDoc.SaveAs strFilePath & strFileName
	wdDoc.Close
	wdApp.Quit
	
	' /* Release memory. */
	Set wdApp = Nothing
	Set wdDoc = Nothing
	Set sldEach = Nothing
	Set sldAll = Nothing
	Set spNotes = Nothing
	Set spNotesPage = Nothing
	Set pptPresentation = Nothing
End Function

' #########################################
' Get file path from a specified full path.
' #########################################
Function GetFilePathFromFullPath(FullPath)
	Dim lngPathSeparatorPosition	' Path separator.
	
	GetFilePathFromFullPath = ""
	lngPathSeparatorPosition = InStrRev(FullPath, "\", -1, 1)
	
	If lngPathSeparatorPosition <> 0 Then GetFilePathFromFullPath = Left(FullPath, lngPathSeparatorPosition)
End Function

' #########################################
' Get file name from a specified full path.
' #########################################
Function GetFileNameFromFullPath(FullPath)
	Dim lngPathSeparatorPosition	' Path separator.
	Dim lngDotPosition				' Dot position.
	Dim strFile						' A full file name.
	
	GetFileNameFromFullPath = ""
	lngPathSeparatorPosition = InStrRev(FullPath, "\", -1, 1)
	
	If lngPathSeparatorPosition <> 0 Then
		strFile = Right(FullPath, Len(FullPath) - lngPathSeparatorPosition)
		lngDotPosition = InStrRev(strFile, ".", -1, 1)
		
		If lngDotPosition <> 0 Then GetFileNameFromFullPath = Left(strFile, lngDotPosition - 1)
	End If
End Function

' ##############################################
' Get file extension from a specified full path.
' ##############################################
Function GetFileExtensionFromFullPath(FullPath)
	Dim lngDotPosition		' Dot position.
	
	GetFileExtensionFromFullPath = ""
	lngDotPosition = InStrRev(FullPath, ".", -1, 1)
	
	If lngDotPosition <> 0 Then GetFileExtensionFromFullPath = Right(FullPath, Len(FullPath) - lngDotPosition)
End Function

' ##################################################################
' Check if the given file extension is PowerPoint presentation file.
' ##################################################################
Function CheckPptFileExtension(FileExtension)
	Dim varArray		' An array contains PowerPoint presentation file extensions.
	Dim varEach			' Each PowerPoint presentation file extension.
	Dim blnIsPptFile	' Whether the file extension is PowerPoint presentation file extension.
	
	blnIsPptFile = False
	
	If FileExtension <> "" Then
		varArray = Array("pptx", "ppt", "pptm", "ppsx", "pps", "ppsm", "odp")
		
		For Each varEach In varArray
			If InStrRev(varEach, FileExtension, -1, 1) <> 0 Then
				blnIsPptFile = True
				Exit For
			End If
		Next
		
	End If
	
	CheckPptFileExtension = blnIsPptFile
End Function
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

Const AS_MSG_DRAGDROP_NOTFOLDER = "Please drag and drop a folder onto this VBScript file. The folder is which you want to save the attachments from the selected Outlook items."
Const AS_MSG_DRAGDROP_MULTIPLE = "Too many files or folders."
Const AS_MSG_ITEMSNOTSELECTED = "Please select an Outlook item at least."
Const AS_MSG_NOATTACHMENTS = "No attachments in the selected Outlook item(s)."
Const AS_ERR_COMPONENT = "There is an error occured, please check the system configuration and make sure Outlook is running correctly."
' The maximum length for a path is 260 characters.
Const MAX_PATH = 260

ExecuteSaving

Sub ExecuteSaving()
	Dim lNum
	
	lNum = SaveAttachmentsFromSelection
	
	If lNum > 0 Then
		MsgBox CStr(lNum) & " attachments are saved successfully.", 64, "Notification"
	Else
		MsgBox AS_MSG_NOATTACHMENTS, 64, "Message"
	End If
End Sub

' ####################################################
' Returns the number of attachements in the selection.
' ####################################################
Function SaveAttachmentsFromSelection()
	Dim fso				' Computer's file system object.
	Dim olApp			' The entire Microsoft Outlook application.
	Dim objItem			' A specific member of a Collection object either by position or by key.
	Dim selItems		' A collection of Outlook item objects in a folder.
	Dim atmt			' A document or link to a document contained in an Outlook item.
	Dim atmtPath		' The full saving path of the attachment.
	Dim atmtFullName	' The full name of an attachment.
	Dim atmtName(1)		' atmtName(0): to save the name; atmtName(1): to save the file extension. They are separated by dot of an attachment file name.
	Dim atmtNameTemp	' To save a temporary attachment file name.
	Dim dotPosition		' The dot position in an attachment name.
	Dim atmts			' A set of Attachment objects that represent the attachments in an Outlook item.
	Dim wsArg			' Command-line parameters.
	Dim lCountEachItem	' The number of attachments in each Outlook item.
	Dim lCountAllItems	' The number of attachments in all Outlook items.
	Dim blnIsSave		' Consider if it is need to save.
	
	blnIsSave = False
	lCountAllItems = 0
	
	On Error Resume Next
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set olApp = GetObject(, "Outlook.Application")
	
	If Err = 0 Then
		wsArg = CGPath(GetFirstArgument)
		
		' /* If the first command-line parameters is a derectory or volume. */
		If fso.FolderExists(wsArg) Then
			Set selItems = olApp.ActiveExplorer.Selection
			
			If Err = 0 Then
				' /* Go through each item in the selection. */
				For Each objItem In selItems
					lCountEachItem = objItem.Attachments.Count
					
					' /* If the current item contains attachments. */
					If lCountEachItem > 0 Then
						
						Set atmts = objItem.Attachments
						
						' /* Go through each attachment in the current item. */
						For Each atmt In atmts
							' Get the full name of the current attachment.
							atmtFullName = atmt.FileName
							
							' Find the dot postion in atmtFullName.
							dotPosition = InStrRev(atmtFullName, ".")
							
							' Get the name.
							atmtName(0) = Left(atmtFullName, dotPosition - 1)
							' Get the file extension.
							atmtName(1) = Right(atmtFullName, Len(atmtFullName) - dotPosition)
							' Get the full saving path of the current attachment.
							atmtPath = wsArg & atmt.FileName
							
							' /* If the length of the saving path is not larger than 260 characters.*/
							If Len(atmtPath) <= MAX_PATH Then
								' True: This attachment can be saved.
								blnIsSave = True
								
								' /* Loop until getting the file name which does not exist in the folder. */
								Do While fso.FileExists(atmtPath)
									atmtNameTemp = atmtName(0) & "_" & SetDateTimeFormat("")
									atmtPath = wsArg & atmtNameTemp & "." & atmtName(1)
									
									' /* If the length of the saving path is over 260 characters.*/
									If Len(atmtPath) > MAX_PATH Then
										lCountEachItem = lCountEachItem - 1
										' False: This attachment cannot be saved.
										blnIsSave = False
										Exit Do
									End If
								Loop
								
								' /* Save the current attachment if it is a valid file name. */
								If blnIsSave Then atmt.SaveAsFile atmtPath
							Else
								lCountEachItem = lCountEachItem - 1
							End If
						Next
					End If
					
					' Count the number of attachments in all Outlook items.
					lCountAllItems = lCountAllItems + lCountEachItem
				Next
			Else
				MsgBox AS_MSG_ITEMSNOTSELECTED, 64, "Message"
				WScript.Quit
			End If
		Else
			MsgBox AS_MSG_DRAGDROP_NOTFOLDER, 48, "Message"
			WScript.Quit
		End If
		
	Else
		MsgBox AS_ERR_COMPONENT, 16, "Error"
		WScript.Quit
	End If
	
	SaveAttachmentsFromSelection = lCountAllItems
	
	' /* Release memory. */
	If Not (fso Is Nothing) Then Set fso = Nothing
	If Not (olApp Is Nothing) Then Set olApp = Nothing
	If Not (objItem Is Nothing) Then Set objItem = Nothing
	If Not (selItems Is Nothing) Then Set selItems = Nothing
	If Not (atmt Is Nothing) Then Set atmt = Nothing
	If Not (atmts Is Nothing) Then Set atmts = Nothing
End Function

' ##########################################
' Get the first parameter from command-line.
' ##########################################
Function GetFirstArgument()
	Dim wsArgNum
	wsArgNum = WScript.Arguments.Count
	
	Select Case wsArgNum
		Case 0
			MsgBox AS_MSG_DRAGDROP_NOTFOLDER, 48 , "Error"
			WScript.Quit
		Case 1
			GetFirstArgument = WScript.Arguments(0)
		Case Else
			MsgBox AS_MSG_DRAGDROP_MULTIPLE, 48 , "Message"
			WScript.Quit
	End Select
End Function

' #################################
' Set the current date time format.
' #################################
Function SetDateTimeFormat(Separator)
	Dim sdtf_Month
	Dim sdtf_Day
	Dim sdtf_Hour
	Dim sdtf_Minute
	Dim sdtf_Second
	Dim sdtf_Millionsecond
	Dim sdtf_TempDate
	Dim sdtf_TempTime
	
	sdtf_TempDate = Date
	sdtf_Month = DatePart("m", sdtf_TempDate)
	sdtf_Day = DatePart("d", sdtf_TempDate)
	
	sdtf_TempTime = Now
	sdtf_Hour = Hour(sdtf_TempTime)
	sdtf_Minute = Minute(sdtf_TempTime)
	sdtf_Millionsecond = Timer * 1000 Mod 1000
	
	If sdtf_Month < 10 Then sdtf_Month = "0" & sdtf_Month
	If sdtf_Day < 10 Then sdtf_Day = "0" & sdtf_Day
	If sdtf_Hour < 10 Then sdtf_Hour = "0" & sdtf_Hour
	If sdtf_Minute < 10 Then sdtf_Minute = "0" & sdtf_Minute
	
	If sdtf_Millionsecond < 100 Then
		If sdtf_Millionsecond < 10 Then
			sdtf_Millionsecond = "00" & sdtf_Millionsecond
		Else
			sdtf_Millionsecond = "0" & sdtf_Millionsecond
		End If
	End If
	
	SetDateTimeFormat = sdtf_Month & Separator & _
						sdtf_Day & Separator & _
						sdtf_Hour & Separator & _
						sdtf_Minute & Separator & _
						sdtf_Millionsecond
End Function

' #####################
' Convert general path.
' #####################
Function CGPath(Path)
	If Right(Path, 1) <> "\" Then Path = Path & "\"
	CGPath = Path
End Function
'=*=*=*=*=*=*=*=*=*=*=*=
' ExportOutlookFolders - Selected PST.vbs
' Coded By Assaf Miron 
' Date : 10/11/06 
'=*=*=*=*=*=*=*=*=*=*=*=
On Error Resume Next 
'======================================================================================================================
' This script Exports all the Mail Object in a Mail Box to MSG files (Default - you can Change the output to any kind you want)
' Exports Calendar meetings, Contacts Cards, Notes, Tasks and Mail objects.
' It builds up your PST Folder Structure and organizes the mail objects by Sender. 
' The script knows to ignore specific folders that you don’t want to export like Calendar or Junk Mail.
' Theoretically it can also export Public Folders but it wasn’t tested
' The script opens a select box in order to select a pst and Export all the mail from it
'======================================================================================================================
'======================================================================================================================
' Consts
'======================================================================================================================
'Mail Formats
Const olTXT = 0
Const olRTF = 1
Const olTemplate = 2
Const olMSG = 3
Const olDoc = 4
Const olHTML = 5
Const olVCard = 6
Const olVCal = 7
'Defualt Mail Folders
Const olFolderDeletedItems = 3
Const olFolderOutbox = 4
Const olFolderSentMail = 5
Const olFolderInbox = 6
Const olFolderCalendar = 9
Const olFolderContacts = 10
Const olFolderJournal = 11
Const olFolderNotes = 12
Const olFolderTasks = 13
Const olFolderDrafts = 16
Const olPublicFoldersAllPublicFolders = 18
Const olFolderJunk = 23
'Exported Folder Path
Const FolderPath = "C:\"
'======================================================================================================================
' Set Objects
'======================================================================================================================
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.Shell")
Set objOutlook = CreateObject("Outlook.Application")
Set objExplorer = CreateObject("InternetExplorer.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
UserProfile = objShell.ExpandEnvironmentStrings("%userprofile%")
'======================================================================================================================
'Open the Users PST File using the Common Dialog box
'It filters out only PST files. you can Add more types by adding |Type Name|*.Type EXT
'======================================================================================================================
Set objDialog = CreateObject("UserAccounts.CommonDialog")
objDialog.Filter = "PST Files|*.pst"
objDialog.InitialDir = UserProfile & "\Local Settings\Application Data\Microsoft\Outlook"
intResult = objDialog.ShowOpen
 
If intResult = 0 Then
    Wscript.Quit
Else
    UsersPST=objDialog.FileName
End If
'======================================================================================================================
'Add the Users Pst File to Outlook using the AddStore method
'Open the pst file using the OpenMAPIFolder function
'======================================================================================================================
objNamespace.Addstore UsersPST
Set oFolder = OpenMAPIFolder("\" & objNameSpace.Folders.GetLast)

strParentFolder = oFolder.Name

'======================================================================================================================
'Subs
'======================================================================================================================
Sub CreateWaitingPage()
'======================================================================================================================
' This Sub creates an HTML waiting page and saves it in the current folder from where the script is running from.
'======================================================================================================================
Set objWaitFile = objFSO.CreateTextFile("Progress.html")
objWaitFile.WriteLine "<html>"
objWaitFile.WriteLine "<head>"
objWaitFile.WriteLine "	<title>Please wait</title>"
objWaitFile.WriteLine "</head>"
objWaitFile.WriteLine ""
objWaitFile.WriteLine "<body bgcolor='buttonface'>"
objWaitFile.WriteLine ""
objWaitFile.WriteLine "<p>"
objWaitFile.WriteLine "<object classid='clsid:35053A22-8589-11D1-B16A-00C0F0283628' id='ProgressBar1' height='20' width='400'>"
objWaitFile.WriteLine "    <param name='Min' value='0'>"
objWaitFile.WriteLine "    <param name='Max' value='100'>"
objWaitFile.WriteLine "    <param name='Orientation' value='0'>"
objWaitFile.WriteLine "    <param name='Scrolling' value='1'>"
objWaitFile.WriteLine "</object>"
objWaitFile.WriteLine "</p>"
objWaitFile.WriteLine ""
objWaitFile.WriteLine "</body>"
objWaitFile.WriteLine "</html>"
objWaitFile.Close

End Sub

Public Sub ExplorerPage()
'======================================================================================================================
' This sub creates an Explorer object, calls the waiting page sub and loads the waiting page positioning it in the center of screen.
'======================================================================================================================
	On Error Resume Next
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colExpItems = objWMIService.ExecQuery("Select * From Win32_DesktopMonitor")
	For Each objExpItem in colExpItems
		intHorizontal = objExpItem.ScreenWidth
		intVertical = objExpItem.ScreenHeight
	Next
	CreateWaitingPage
	wscript.sleep 200
	strCurrentFolder = objShell.CurrentDirectory
	
	objExplorer.Navigate "file:///" & strCurrentFolder & "\Progress.html"   
	objExplorer.ToolBar = 0
	objExplorer.StatusBar = 0
	objExplorer.Width = 450
	objExplorer.Height = 90 
	objExplorer.Left = (intHorizontal - 450) / 2
	objExplorer.Top = (intVertical - 90) / 2
	objExplorer.Visible = 1             
	
	Do While (objExplorer.Busy)
		Wscript.Sleep 200
	Loop 

End Sub

Public Sub CreateOLKFolder(strFolderName,strFolderPath)
'======================================================================================================================
' This Sub receives a folder name and the path of the folder structure.
' at first it zero's out the progress bar
' Checks if this folder has any sub folders
' defines a new collection of outlook folders and items
' and creates a folder for every outlook folder and saves all of its outlook objects in it according to the path given earlier
'======================================================================================================================
objExplorer.Document.Body.All.ProgressBar1.Value = 0
		Set colSubFolders = strFolderName.Folders
		If colSubFolders.Count => 1 Then
			For Each olkSubFolder in colSubFolders
				If NotExluded(olkSubFolder)=-1 Then
					strSubFolderPath = strFolderPath & "\" & olkSubFolder.name
					If Not objFSO.FolderExists(strSubFolderPath) then
						Set objFolder = objFSO.CreateFolder(strSubFolderPath)
					End If
			
					Set SubcolItems = olkSubFolder.Items
					For Each objItem in SubcolItems
						intItems = SubcolItems.Count ^ 2 ' This was set only for visual Purposes, to keep the progress bar running
						intIncrement = 100/intItems
						olkSubject = objItem.Subject
						olkSender = objItem.SenderEmailAddress
						olkSubject = ElegalSings(olkSubject)
'======================================================================================================================
' Organizes all the mail objects by Sender Email address folders
' Each sender has a folder.
' At the end of this Sub it calls itself agian to check all sub folders of this Foler if any.
'======================================================================================================================
						If Not objFSO.FolderExists(strSubFolderPath & "\" & olkSender) then
							Set objFolder = objFSO.CreateFolder(strSubFolderPath & "\" & olkSender)
						End If
						objItem.SaveAs objFolder & "\" & olkSubject & ".msg", olMSG
						objExplorer.Document.Body.All.ProgressBar1.Value = objExplorer.Document.Body.All.ProgressBar1.Value + intIncrement
					Next
					CreateOLKFolder olkSubFolder, strSubFolderPath
					End If
				Next
			End If
End Sub

'======================================================================================================================
'Functions
'======================================================================================================================
Public Function OpenMAPIFolder(ByVal strPath)
'======================================================================================================================
'Function to get Folder Path 
'I took this Function from the internet (google groups) and modified it a little so it would work here
'http://groups.google.com.fj/group/microsoft.public.office.developer.outlook.vba/browse_thread/thread/d34ebaa8915535ce/a04f656ba89d6631?hl=en
'======================================================================================================================
    On Error Resume Next 
    If Left(strPath, Len("\")) = "\" Then 
        strPath = Mid(strPath, Len("\") + 1) 
    Else 
        Set objFldr = objOutlook.ActiveExplorer.CurrentFolder 
    End If 
    While strPath <> "" 
        i = InStr(strPath, "\") 
        If i Then 
            strDir = Left(strPath, i - 1) 
            strPath = Mid(strPath, i + Len("\")) 
        Else 
            strDir = strPath 
            strPath = "" 
        End If 
        If objFldr Is Nothing Then 
            Set objFldr = objOutlook.GetNamespace("MAPI").Folders(strDir) 
            On Error GoTo 0 
        Else 
            Set objFldr = objFldr.Folders(strDir) 
        End If 
    Wend 
    Set OpenMAPIFolder = objFldr 
       
End Function 

Public Function ElegalSings(olkSubject)
'======================================================================================================================
' To Prevent Stoping the Script on Elegal sings in the Subject, replace all the sings to spaces
' If the subject is empty the write - "Mail Message From - " & Sender Email Address
'======================================================================================================================
	olkSubject = Replace(olkSubject, "\", " ")
	olkSubject = Replace(olkSubject, "/", " ")
	olkSubject = Replace(olkSubject, ":", " ")
	olkSubject = Replace(olkSubject, "?", " ")
	olkSubject = Replace(olkSubject, ">", " ")
	olkSubject = Replace(olkSubject, "<", " ")
	olkSubject = Replace(olkSubject, "|", " ")
	olkSubject = Replace(olkSubject, "*", " ")
	olkSubject = Replace(olkSubject, chr(34), " ") ' chr(34) = "
					
	if olkSubject = "" then
		olkSubject = "Mail Message From - " & olkSender
	End If
	ElegalSings = olkSubject
End Function

Public Function NotExluded(olkFolder)
'======================================================================================================================
' This function defines wich outlook folders you want to Exclude from exporting
' It recives the outlook folder name and if it is Not Exluded, function returns False - Else returns True
' You can mark out any folder you want to include - simply delete its If..End If or mark ' at the begining
'======================================================================================================================
DeletedItems = objNamespace.GetDefaultFolder(olFolderDeletedItems).name
Outbox = objNamespace.GetDefaultFolder(olFolderOutbox).name
SentMail = objNamespace.GetDefaultFolder(olFolderSentMail).name
Inbox = objNamespace.GetDefaultFolder(olFolderInbox).name
Calendar = objNamespace.GetDefaultFolder(olFolderCalendar).name
Contacts = objNamespace.GetDefaultFolder(olFolderContacts).name
Journal = objNamespace.GetDefaultFolder(olFolderJournal).name
Notes = objNamespace.GetDefaultFolder(olFolderNotes).name
Tasks = objNamespace.GetDefaultFolder(olFolderTasks).name
Drafts = objNamespace.GetDefaultFolder(olFolderDrafts).name
'AllPublicFolders = objNamespace.GetDefaultFolder(olFolderAllPublicFolders).name
Junk = objNamespace.GetDefaultFolder(olFolderJunk).name

If olkFolder = DeletedItems Then
	NotExluded = False
	Exit Function
End If

If olkFolder = Outbox Then
	NotExluded = False
	Exit Function
End If

If olkFolder = SentMail Then
	NotExluded = False
	Exit Function
End If

If olkFolder = Inbox Then
	NotExluded = False
	Exit Function
End If

If olkFolder = Calendar Then
	NotExluded = False
	Exit Function
End If

If olkFolder = Contacts Then
	NotExluded = False
	Exit Function
End If

If olkFolder = Journal Then
	NotExluded = False
	Exit Function
End If

If olkFolder = Notes Then
	NotExluded = False
	Exit Function
End If

If olkFolder = Tasks Then
	NotExluded = False
	Exit Function
End If

If olkFolder = Drafts Then
	NotExluded = False
	Exit Function
End If

If olkFolder = Junk Then
	NotExluded = False
	Exit Function
End If

NotExluded = True
	
End Function

'======================================================================================================================
'Code 
'======================================================================================================================
ExplorerPage ' Calls the Explorer page
'======================================================================================================================
' Checks if the Parent folder of the PST exists if not - Creates it
' Define objMailbox as the PST root folder and gather its sub folders in colFolders
'======================================================================================================================
	If Not objFSO.FolderExists(FolderPath & strParentFolder) then
		Set objFolder = objFSO.CreateFolder(FolderPath & strParentFolder)
	End If
	Set objMailbox = objNamespace.Folders(strParentFolder)
	Set colFolders = objMailbox.Folders

	For Each olkFolder in colFolders
		If NotExluded(olkFolder)=-1 Then
			strFolderPath =FolderPath & strParentFolder & "\" & olkFolder.name
			If Not objFSO.FolderExists(strFolderPath) then
				Set objFolder = objFSO.CreateFolder(strFolderPath)
			End If
			Set colItems = olkFolder.Items
			For Each objItem in colItems
				intItems = colItems.Count ^ 2 ' This was set only for visual Purposes, to keep the progress bar running
				intIncrement = 100/intItems
				olkSubject = objItem.Subject
				olkSender = objItem.SenderEmailAddress
				olkSubject = ElegalSings(olkSubject)
'======================================================================================================================
' Organizes all the mail objects by Sender Email address folders
' Each sender has a folder.
' At the end of this Sub it calls itself agian to check all sub folders of this Foler if any.
'======================================================================================================================
				If Not objFSO.FolderExists(strFolderPath & "\" & olkSender) then
					Set objFolder = objFSO.CreateFolder(strFolderPath & "\" & olkSender)
				End If
				objItem.SaveAs objFolder & "\" & olkSubject & ".msg", olMSG
				objExplorer.Document.Body.All.ProgressBar1.Value = objExplorer.Document.Body.All.ProgressBar1.Value + intIncrement
			Next
			CreateOLKFolder olkFolder,strFolderPath
			objExplorer.Document.Body.All.ProgressBar1.Value = 90
			Wscript.Sleep 500
			objExplorer.Document.Body.All.ProgressBar1.Value = 100
		End If
	Next
'======================================================================================================================
' CleanUp
' Exit the Explorer page that shows the progress bar
' Delete the progress.html File
'======================================================================================================================
objExplorer.Quit
strCurrentFolder = objShell.CurrentDirectory
objFSO.DeleteFile(strCurrentFolder & "\Progress.html" )	
wscript.echo "Done!"
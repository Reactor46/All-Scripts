'==================================================================================
' Script: 	QueueStatsCookdown.vbs
' Date:		5/25/12	
' Author: 	Brian Wren, Microsoft Corporation
' Purpose:	Collects statistics for the entire set of queue folders for StoreApp sample application.
'			Script supports cookdown because it doesn't require any arguments that will vary for each target instance.
'==================================================================================

SetLocale("en-us")

'Constants used for event logging
Const SCRIPT_NAME			= "QueueStatsCookdown.vbs"
Const EVENT_LEVEL_ERROR 	= 1
Const EVENT_LEVEL_WARNING 	= 2
Const EVENT_LEVEL_INFO 		= 4

Const SCRIPT_STARTED		= 831
Const PROPERTYBAG_CREATED	= 832
Const SCRIPT_ENDED			= 835

'Setup variables sent in through script arguments
sTopFolder = WScript.Arguments(0)		'Path of the top level folder where queue folders are located.
bDebug = CBool(WScript.Arguments(1))	'If true, information events are loggged..

'Start by setting up API object.
Set oAPI = CreateObject("MOM.ScriptAPI")
sMessage =	"Top Folder: " & sTopFolder
Call LogDebugEvent(SCRIPT_STARTED,sMessage)

'Get the FileSystemObject and the top level folder.
Set fso = CreateObject("Scripting.FileSystemObject")
Set oFolder = fso.GetFolder(sTopFolder)

'Enumerate each folder under the top level.
For Each oSubFolder in oFolder.SubFolders
	
	'Create a property bag for each folder.
	Set oBag = oAPI.CreatePropertyBag()
	
	'Get the statistics we're interested in for the current subfolder.
	iCount = oSubFolder.Files.Count
	iSize = 0
	iOldestFile = 0
	For Each oFile In oSubFolder.Files
		iSize = iSize + oFile.Size
		iAgeInMinutes = DateDiff("n",oFile.DateCreated,Now)
		If iAgeInMinutes > iOldestFile Then
			iOldestFile = iAgeInMinutes
		End If
	Next 
	
	sMessage =	"Property bag created" & VbCrLf & _
				"StoreCode: " & oSubFolder.Name & VbCrLf & _
				"FileCount: " & iCount & VbCrLf & _
				"OldestFile: " & iOldestFile & VbCrLf & _
				"TotalSize: " & iSize
	Call LogDebugEvent(PROPERTYBAG_CREATED,sMessage)
	
	'Put the gathered statistics into the property bag.  
	'Includes a value for the folder name so that we can tell which folder the data is from.
	Call oBag.AddValue("StoreCode",oSubFolder.Name)
	Call oBag.AddValue("FileCount",iCount)
	Call oBag.AddValue("OldestFile",iOldestFile)
	Call oBag.AddValue("TotalSize",iSize)
	Call oAPI.AddItem(oBag)
Next

'Return all property bags.
oAPI.ReturnItems()

Call LogDebugEvent (SCRIPT_ENDED,"Script ended.")

'==================================================================================
' Sub:		LogDebugEvent
' Purpose:	Logs an informational event to the Operations Manager event log 
'			only if Debug argument is true
'==================================================================================
Sub LogDebugEvent(EventNo,Message)

	Message = VbCrLf & Message
	If bDebug = True Then
    	Call oAPI.LogScriptEvent(SCRIPT_NAME,EventNo,EVENT_LEVEL_INFO,Message)
	End If
	
End Sub   
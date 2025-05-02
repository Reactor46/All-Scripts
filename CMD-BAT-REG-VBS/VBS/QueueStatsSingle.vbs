'==================================================================================
' Script: 	QueueStatsSingle.vbs
' Date:	    5/25/12	
' Author: 	Brian Wren, Microsoft Corporation
' Purpose:	Collects statistics for a single queue folder for StoreApp sample application.
'		    Script does not support cookdown because it requires the folder path as an argument.
'		    Since we expect multiple instances of queue folders on a particular agent, 
'			each instance will have its own instance of the script.
'==================================================================================

SetLocale("en-us")

'Constants used for event logging
Const SCRIPT_NAME			= "QueueStatsSingle.vbs"
Const EVENT_LEVEL_ERROR 	= 1
Const EVENT_LEVEL_WARNING 	= 2
Const EVENT_LEVEL_INFO 		= 4

Const SCRIPT_STARTED		= 821
Const PROPERTYBAG_CREATED	= 822
Const SCRIPT_ENDED		= 825

'Setup variables sent in through script arguments
sFolder = WScript.Arguments(0)			'Path of the folder we'll be collecting statistics for.
bDebug = CBool(WScript.Arguments(1))	'If true, information events are loggged..

'Start by setting up API object.
Set oAPI = CreateObject("MOM.ScriptAPI")

'Log a message that script is starting only if Debug argument is True
sMessage =	"Script started" & VbCrLf & _
		"Folder: " & sFolder
Call LogDebugEvent(SCRIPT_STARTED,sMessage)

'Create a property bag.
Set oBag = oAPI.CreatePropertyBag()

'Get the FileSystemObject and the folder.
Set fso = CreateObject("Scripting.FileSystemObject")
Set oFolder = fso.GetFolder(sFolder)

'Get the statistics we're interested in for the specified folder.	
iCount = oFolder.Files.Count
iSize = 0
iOldestFile = 0
For Each oFile In oFolder.Files
	iSize = iSize + oFile.Size
	iAgeInMinutes = DateDiff("n",oFile.DateCreated,Now)
	If iAgeInMinutes > iOldestFile Then
		iOldestFile = iAgeInMinutes
	End If
Next 
sMessage =	"Property bag created" & VbCrLf & _
		"FileCount: " & iCount & VbCrLf & _
		"OldestFile: " & iOldestFile & VbCrLf & _
		"TotalSize: " & iSize
Call LogDebugEvent(PROPERTYBAG_CREATED,sMessage)

'Put the gathered statistics into the property bag.
Call oBag.AddValue("FileCount",iCount)
Call oBag.AddValue("OldestFile",iOldestFile)
Call oBag.AddValue("TotalSize",iSize)

'Return the property bag.
Call oAPI.Return(oBag)

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
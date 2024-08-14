'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' AUTHOR  : Marius @ Hican - http://www.hican.nl - @hicannl 
' DATE    : 25-04-2012
' COMMENT : This script auto creates (removal) 
'           Collections regarding SMS / SCCM based on an 
'           input csv file. Including Logging.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
 
Const ForWriting = 2
Const ForReading = 1
Const ForAppending = 8
 
Dim objWshShell, objFso, objFile, objEnv, objLogFileFSO, objLogFile
Dim swbemLocator, swbemconnection, providerLoc, Location
Dim sConnect, sCurPath, sInputFile, sOutputFile, sNextLine, sSplit
Dim sParent, sAppName, sAll, sC2R, sInstalls, sUsers, sUsrName, sComment
Dim sColCheck, Collection, objCollection, newCollection, collectionPath
Dim sCollectionID, objContainerNode, Container, ParentFolderID, sResourceID
Dim newCollectionRelation, Token, objNewCollection, sColInput, sLimitID
Dim newQueryRule, newCollectionRule, sQueryC2R, sQueryInstalls, sQueryUsers
Dim sGroupName, sProductID
 
sConnect        = "<SERVER_NAME>"
sCurPath        = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))
sInputFile      = sCurPath + Chr(92) + "input.csv"
sOutputFile     = sCurPath + Chr(92) + "output.csv"
 
Set objWshShell = WScript.CreateObject("WScript.Shell")
Set objFso      = CreateObject("Scripting.FileSystemObject")
Set objFile     = objFso.OpenTextFile(sInputFile, ForReading)
Set objEnv      = objWshShell.Environment("process")
 
'Try to make a connection with the SCCM server.
ConnectToSCCM
 
OpenLogFile()
WriteToLog "-------------------------------------------------------------------" & VbCrLf
WriteToLog "START LOG  (" + FormatDateTime(Now(), vbGeneralDate)
 
Do Until objFile.AtEndOfStream
  sNextLine  = objFile.Readline
  sSplit     = Split(sNextLine, ";")
  sParent    = sSplit(0)
  sAppName   = sSplit(1)
  sProductID = sSplit(2)
  sGroupName = sSplit(3)
 
  sAll       = "All_" + sAppName
  sC2R       = "C2R_" + sAppName
  sInstalls  = "Installs_" + sAppName
  sUsers     = "Users_" + sAppName
 
  sQueryC2R      = "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System INNER JOIN SMS_R_User ON SMS_R_System.LastLogonUserName = SMS_R_User.UserName WHERE (SMS_R_System.Name NOT IN (SELECT SMS_R_System.Name FROM SMS_R_User INNER JOIN SMS_R_System ON SMS_R_User.Username = SMS_R_System.LastLogonUserName WHERE (SMS_R_User.UserGroupName = 'Hican.net\\" + sGroupName + "'))) AND (SMS_R_System.Name IN (SELECT SMS_R_System.Name FROM SMS_R_System AS SMS_R_System INNER JOIN SMS_G_System_ADD_REMOVE_PROGRAMS ON SMS_R_System.ResourceID = SMS_G_System_ADD_REMOVE_PROGRAMS.ResourceID WHERE (SMS_G_System_ADD_REMOVE_PROGRAMS.ProdID = '" + sProductID + "')))"
  sQueryInstalls = "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_ADD_REMOVE_PROGRAMS on SMS_G_System_ADD_REMOVE_PROGRAMS.ResourceID = SMS_R_System.ResourceId where SMS_G_System_ADD_REMOVE_PROGRAMS.ProdID = '" + sProductID + "'"
  sQueryUsers    = "select SMS_R_USER.ResourceID,SMS_R_USER.ResourceType,SMS_R_USER.Name,SMS_R_USER.UniqueUserName,SMS_R_USER.WindowsNTDomain from SMS_R_USER WHERE (SMS_R_USER.USERGROUPNAME = 'Hican.net\\" + sGroupName + "')"
 
  sUsrName  = objEnv("USERNAME")
  sComment  = "Created via script by " + sUsrName + " on " + Replace(FormatDateTime(Now(), vbShortDate), "/", "-")
  sColCheck = ""
 
  If (Left(sNextLine, 1) = "*") Then
    WriteToLog "LINE: " & sNextLine & " DISABLED FOR PROCESSING..."
  Else
    CreateCollections(sAll)
    CreateCollections(sC2R)
    CreateCollections(sInstalls)
    CreateCollections(sUsers)
  End If
Loop
 
objFile.Close
 
WriteToLog "-------------------------------------------------------------------" & VbCrLf
CloseLogFile()
Wscript.Echo "Finished creating the Collections!"
Wscript.Quit
 
'''''''''''''''''''''''''''''
'         FUNCTIONS         '
'''''''''''''''''''''''''''''
Function ConnectToSCCM
On Error Resume Next
Set swbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set swbemconnection = swbemLocator.ConnectServer(sConnect, "root\sms")
Set providerLoc = swbemconnection.InstancesOf("SMS_ProviderLocation")
 
For Each Location In providerLoc
  If location.ProviderForLocalSite = True Then
    Set swbemconnection = swbemLocator.ConnectServer(Location.Machine, "root\sms\site_" + Location.SiteCode)
    Exit For
  End If
Next
If Err.Number <> 0 Then
  Wscript.echo "Unable to connect to the SCCM provider. " & _
               "Check the connection and / or the settings in the script!" & _
               "The script will be stopped now."
  WScript.Quit
End If
On Error GoTo 0
End Function
 
Sub OpenLogFile()  	
Set objLogFileFSO = CreateObject("Scripting.FileSystemObject")
If objLogFileFSO.FileExists(sOutputFile) Then
  Set objLogFile = objLogFileFSO.OpenTextFile(sOutputFile, ForAppending)
Else
  Set objLogFile = objLogFileFSO.CreateTextFile(sOutputFile)
End If	
End Sub
 
Sub CloseLogFile()
objLogFile.Close	
Set objLogfileFSO = Nothing
End Sub
 
Function WriteToLog(sLogMessage)
  objLogFile.WriteLine(sLogMessage)
End Function
 
Function ConvertToWMIdate(sDate)
'Attempts to convert the date into a WMI date-time!
Dim sYear, sMonth, sDay, sHour, sMinute
 
sYear = year(sDate)
sMonth = month(sDate)
sDay = day(sDate)
sHour = hour(sDate)
sMinute = minute(sDate)
 
If len(sMonth) = 1 Then
  sMonth = "0" & sMonth
End If
If len(sDay) = 1 Then
  sDay = "0" & sDay
End If
If len(sHour) = 1 Then
  sHour = "0" & sHour
End If
If len(sMinute) = 1 Then
  sMinute = "0" & sMinute
End If
 
ConvertToWMIdate = sYear & sMonth & sDay & sHour & sMinute & "00.000000+***"
End Function
 
 
Function CreateCollections(sColInput)
'Attempts to check if the collection already exists. If it does, 
'the script will skip this creation!
Set Collection = swbemconnection.ExecQuery ("select * from SMS_Collection where Name='" & sParent & "'")
For Each objCollection In Collection
  sColCheck = "This collection exists with the collection ID of: " & objCollection.CollectionID
Next
 
If sColCheck = "" Then
  WScript.Echo "ParentFolder " + sParent + " is not present in Collections - " & _
               "Management Collections - Special Collections - SMS Cleanup " & _ 
               ", please create that folder first. The script will be stopped now."
  WScript.Quit
Else
  'Attempts to create the new collections.
  Set newCollection = swbemconnection.Get("SMS_Collection").SpawnInstance_()
 
  newCollection.Name = sColInput
  newCollection.OwnedByThisSite = True
  newCollection.Comment = sComment
  Set collectionPath = newCollection.Put_
 
  'Attempts to obtain the collection ID of the 
  'newly created collection!
  Set Collection = swbemconnection.ExecQuery ("select * from SMS_Collection where Name='" & sColInput & "'")
  For Each objCollection in Collection
    sCollectionID = objCollection.CollectionID
  Next
 
  If sCollectionID = "" Then
    WScript.Echo "Unable to obtain a collection ID for the newly created collection."
    WScript.Quit
  Else
    'WScript.Echo sCollectionID
  End If
 
  If sColInput = sAll Then
    Set objContainerNode = swbemconnection.ExecQuery("select * from SMS_Collection where Name='" & sParent & "'")
  Else
    Set objContainerNode = swbemconnection.ExecQuery("select * from SMS_Collection where Name='" & sAll & "'")
  End If
 
  'ParentFolderID = ""
  For Each Container In objContainerNode
    If Container.name = sParent Then
      ParentFolderID = Container.CollectionID
    Elseif Container.name = sAll Then
      ParentFolderID = Container.CollectionID
    Else
      'WScript.Echo ParentFolderID
    End If
  Next
 
  'Attempts to move the newly created collection into the 
  'desired parent collection.
  Set newCollectionRelation = swbemconnection.Get("SMS_CollectToSubCollect").SpawnInstance_()
  newCollectionRelation.parentCollectionID = ParentFolderID
  newCollectionRelation.subCollectionID = sCollectionID
  newCollectionRelation.Put_
 
  'Attempts to create and add the query rule group
  'to the collection!
  If sColInput = sC2R OR sColInput = sInstalls OR sColInput = sUsers Then
    Set Collection = swbemconnection.ExecQuery ("select * from SMS_Collection where Name='All Managed Workstations'")
    For Each objCollection in Collection
      sLimitID = objCollection.CollectionID
    Next
 
    Set newQueryRule = swbemconnection.Get("SMS_CollectionRuleQuery").SpawnInstance_()
    If sColInput = sC2R Then
      newQueryRule.QueryExpression     = sQueryC2R
      newQueryRule.LimitToCollectionID = sLimitID
    Elseif sColInput = sInstalls Then
      newQueryRule.QueryExpression     = sQueryInstalls
      newQueryRule.LimitToCollectionID = sLimitID
    Elseif sColInput = sUsers Then
      newQueryRule.QueryExpression     = sQueryUsers
    End If
    newQueryRule.RuleName = sColInput
 
    'Add the new query rule to a variable.
    Set newCollectionRule = newQueryRule
    'Get the collection.
    Set newCollection = swbemconnection.Get(collectionPath.RelPath)
    'Add the rules to the collection.
    newCollection.AddMembershipRule newCollectionRule
    newCollection.RequestRefresh False 
  End If
 
  'Attempts to set the membership update schedule on the 
  'collection (weekly recurrance)!
  Set Token = swbemconnection.Get("SMS_ST_RecurWeekly")
  Token.StartTime = ConvertToWMIdate(Now())
 
  Set objNewCollection = swbemconnection.Get ("SMS_Collection.CollectionID='" & sCollectionID & "'")
  objNewCollection.RefreshSchedule = Array(Token)
  objNewCollection.RefreshType = 2
  objNewCollection.Put_
 
  WriteToLog "CREATED: " & sColInput
End If
End Function
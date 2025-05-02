const SUBSTRING      = 1             ' Substring
const IGNORECASE     = &H00010000    ' Ignore case
Const ACTION_MOVE 	   = 1
const B_NEZ                = 2
const L_OR            	   = 3
Const REL_EQ                = 7
const MSG_ATTACH            = 2
Const PR_Transport_Headers = &H007D001E

servername = "servername"
mailboxname = "user"

Set objSession   = CreateObject("MAPI.Session")

objSession.Logon "","",false,true,true,true,servername & vbLF & mailboxname
Set objRules     = CreateObject("MSExchange.Rules")
objRules.Folder  = objSession.Inbox
Set objInbox = objSession.Inbox

Set CdoInfoStore = objSession.GetInfoStore
Set CdoFolderRoot = CdoInfoStore.RootFolder
Set CdoFolders = CdoFolderRoot.Folders

bFound = False
Set CdoFolder = CdoFolders.GetFirst
Do While (Not bFound) And Not (CdoFolder Is Nothing)
    If CdoFolder.Name = "Junk E-mail" Then
       bFound = True
    Else
       Set CdoFolder = CdoFolders.GetNext
    End If
Loop
Set ActionFolder = CdoFolder


Set importPropVal      = CreateObject("MSExchange.PropertyValue")
importPropVal.Tag      = PR_Transport_Headers
importPropVal.Value    = " attachment;"

Set importPropCond         = CreateObject("MSExchange.ContentCondition")
importPropCond.PropertyType = PR_Transport_Headers
importPropCond.Operator    =  SUBSTRING + IGNORECASE
importPropCond.Value       = importPropVal


Set importPropVal1      = CreateObject("MSExchange.PropertyValue")
importPropVal1.Tag      = PR_Transport_Headers
importPropVal1.Value    = "image/gif"

Set importPropCond1         = CreateObject("MSExchange.ContentCondition")
importPropCond1.PropertyType = PR_Transport_Headers
importPropCond1.Operator    =  SUBSTRING + IGNORECASE
importPropCond1.Value       = importPropVal1


Set logPropCond1      = CreateObject("MSExchange.LogicalCondition")
logPropCond1.Operator =  3
logPropCond1.Add importPropCond

Set logPropCond      = CreateObject("MSExchange.LogicalCondition")
logPropCond.Operator =  1
logPropCond.Add importPropCond1
logPropCond.Add logPropCond1


' Create action
Set objAction        = CreateObject("MSExchange.Action")
objAction.ActionType = ACTION_MOVE 
objAction.Arg = ActionFolder

' Create new rule
Set objRule   = CreateObject("MSExchange.Rule")
objRule.Name  = "Gif Image Move Rule"

' Add action and assign condition
objRule.Actions.Add   , objAction
objRule.Condition   = logPropCond 

' Add rule and update
objRules.Add  , objRule
objRules.Update

' Log off and cleanup
objSession.Logoff

Set objRules       = Nothing
Set objSession     = Nothing
Set importProp     = Nothing
Set importPropVal  = Nothing
Set objAction      = Nothing
Set objRule        = Nothing
const REL_EQ               = 7
Const ACTION_MOVE 	   = 1
const CdoPR_CONTENT_FILTER_SCL     = &H40760003
const SCL_VAL       = 5

servername = "servername"
mailboxname = "mailboxalias"

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
importPropVal.Tag      = CdoPR_CONTENT_FILTER_SCL
importPropVal.Value    = SCL_VAL

Set importPropCond         = CreateObject("MSExchange.PropertyCondition")
importPropCond.PropertyTag = CdoPR_CONTENT_FILTER_SCL
importPropCond.Operator    = REL_EQ
importPropCond.Value       = importPropVal

' Create action
Set objAction        = CreateObject("MSExchange.Action")
objAction.ActionType = ACTION_MOVE
objAction.Arg = ActionFolder

' Create new rule
Set objRule   = CreateObject("MSExchange.Rule")
objRule.Name  = "SCL VAL Move"

' Add action and assign condition
objRule.Actions.Add   , objAction
objRule.Condition   = importPropCond

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
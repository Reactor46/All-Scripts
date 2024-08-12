const SUBSTRING      = 1             ' Substring
const IGNORECASE     = &H00010000    ' Ignore case
const PR_SENDER_EMAIL_ADDRESS     = &H0C1F001E
Const ACTION_REPLY = 4


servername = "servername"
mailboxname = "mailbox"

Set objSession   = CreateObject("MAPI.Session")

objSession.Logon "","",false,true,true,true,servername & vbLF & mailboxname
Set objRules     = CreateObject("MSExchange.Rules")
objRules.Folder  = objSession.Inbox
Set objInbox = objSession.Inbox

Set CdoInfoStore = objSession.GetInfoStore
Set CdoFolderRoot = CdoInfoStore.RootFolder
Set CdoFolders = CdoFolderRoot.Folders

Set importPropVal      = CreateObject("MSExchange.PropertyValue")
importPropVal.Tag      = PR_SENDER_EMAIL_ADDRESS
importPropVal.Value    = "@yahoo.com"

Set importPropCond         = CreateObject("MSExchange.ContentCondition")
importPropCond.PropertyType = PR_SENDER_EMAIL_ADDRESS
importPropCond.Operator    =  SUBSTRING + IGNORECASE
importPropCond.Value       = importPropVal

' Create reply message and store in HiddenMessages collection.
Set objReplyMsg = objInbox.HiddenMessages.Add

' Set reply message properties.
objReplyMsg.Type = "IPM.Note.Rules.ReplyTemplate.Microsoft"
objReplyMsg.Text = "Im Sorry this Mailbox isn't currently maned for after hour enquires please contact 333-333-33"
objReplyMsg.Update


' Create action
Set objAction        = CreateObject("MSExchange.Action")
objAction.ActionType = ACTION_REPLY
objAction.Arg = Array(objReplyMsg.ID,objReplyMsg.FolderID)

' Create new rule
Set objRule   = CreateObject("MSExchange.Rule")
objRule.Name  = "Domain Reply Rule"

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
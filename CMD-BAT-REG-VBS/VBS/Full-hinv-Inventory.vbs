'Declare Variables 
On Error Resume Next 
Set sho = CreateObject("WScript.Shell") 
strSystemRoot = sho.expandenvironmentstrings("%SystemRoot%") 
strCurrentDir = Left(Wscript.ScriptFullName, (InstrRev(Wscript.ScriptFullName, "\") -1)) 
' Get a connection to the "root\ccm\invagt" namespace (where the Inventory agent lives) 
Dim oLocator 
Set oLocator = CreateObject("WbemScripting.SWbemLocator") 
Dim oServices 
Set oServices = oLocator.ConnectServer( , "root\ccm\invagt") 
'Reset SMS Hardware Inventory Action to force a full HW Inventory Action 
sInventoryActionID = "{00000000-0000-0000-0000-000000000101}" 
' Delete the specified InventoryActionStatus instance 
oServices.Delete "InventoryActionStatus.InventoryActionID=""" & sInventoryActionID & """" 
'Pause 3 seconds To allow the action to complete. 
wscript.sleep 3000 
'Run a SMS Hardware Inventory 
Set cpApplet = CreateObject("CPAPPLET.CPAppletMgr") 
Set actions = cpApplet.GetClientActions 
For Each action In actions 
If Instr(action.Name,"Hardware Inventory Collection") > 0 Then 
action.PerformAction 
End If 
Next
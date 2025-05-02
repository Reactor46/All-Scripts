Const FL_FORCE_CREATE_NAMESPACE = 4 
 
strComputer = "." 
 
Set wDate = CreateObject("WbemScripting.SWbemDateTime") 
Set locator = CreateObject("WbemScripting.SWbemLocator") 
Set connection = locator.ConnectServer( strComputer, "root\rsop", null, null, null, null, 0, null) 
Set provider = connection.Get("RsopLoggingModeProvider") 
provider.RsopCreateSession FL_FORCE_CREATE_NAMESPACE, Null, namespaceLocation, hResult, eInfo 
 
Set rsopProv = locator.ConnectServer _ 
    (strComputer, namespaceLocation & "\User", null, null, Null, Null, 0 , Null) 
 
WScript.Echo "User Scope of Management Logging" 
 
Set colItems = rsopProv.ExecQuery("Select * from RSOP_SOM") 
 
For Each objItem in colItems 
    WScript.Echo String(50, "=") 
    Wscript.Echo "ID: " & objItem.ID 
    Wscript.Echo "Blocked: " & objItem.Blocked 
    Wscript.Echo "Blocking: " & objItem.Blocking 
    Select Case objItem.Reason 
        Case 1 
                        WScript.Echo "Mode: Normal" 
        Case 2 
                        WScript.Echo "Mode: Loopback" 
    End Select    
    Wscript.Echo "SOM Order: " & objItem.SOMOrder 
    Select Case objItem.Type 
        Case 1 
                        WScript.Echo "Scope:  Local" 
        Case 2 
                        WScript.Echo "Scope:  Site" 
        Case 3 
                        WScript.Echo "Scope: Domain" 
        Case 4 
                        WScript.Echo "Organizational Unit" 
        End Select  
Next 
 
provider.RsopDeleteSession namespaceLocation, hResult 
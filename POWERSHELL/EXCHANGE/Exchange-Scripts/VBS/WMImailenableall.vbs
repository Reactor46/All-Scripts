On Error Resume Next
Dim cComputerName
Const cWMINameSpace = "root/MicrosoftExchangeV2"
Const cWMIInstance = "Exchange_PublicFolder"
cComputerName = "Servername"

strWinMgmts = "winmgmts:{impersonationLevel=impersonate}!//"& _
cComputerName&"/"&cWMINameSpace
Set objWMIExchange =  GetObject(strWinMgmts)
If Err.Number <> 0 Then
  WScript.Echo "ERROR: Unable to connect to the WMI namespace."
Else
  Set listExchange_PublicFolders = objWMIExchange.InstancesOf(cWMIInstance)
  If (listExchange_PublicFolders.count > 0) Then
  For Each objExchange_PublicFolder in listExchange_PublicFolders    
       WScript.echo objExchange_PublicFolder.Name & " " & objExchange_PublicFolder.path & " " & objExchange_PublicFolder.IsMailEnabled & " " & objExchange_PublicFolder.IsSearchFolder
       if objExchange_PublicFolder.IsMailEnabled <> true then
		objExchange_PublicFolder.IsMailEnabled = true 
       		objExchange_PublicFolder.Put_()
		Wscript.echo "Mail Enabled Public Folder" & objExchange_PublicFolder.Name
       end if 
    Next
  Else
    WScript.Echo "WARNING: No Exchange_PublicFolder instances were returned."
  End If
End If


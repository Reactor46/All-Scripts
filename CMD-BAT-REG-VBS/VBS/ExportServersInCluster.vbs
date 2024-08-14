On Error Resume Next 
Dim objFSO,objshell,envUSER,TextPath,strComputer	
'Create "Scripting.FileSystemObject" object 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
'Create "Wscript.shell" object 
Set objshell = CreateObject("wscript.shell")
'Get the current username
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\mscluster")
Set colItems = objWMIService.ExecQuery("Select * from MSCluster_Node")
If Err.Number = 0 Then 
	envUSER = objshell.expandEnvironmentStrings("%USERPROFILE%")
	TextPath = envUSER & "\\Desktop\ClusterNodes.txt"
	Set Txt = objFSO.CreateTextFile(TextPath)
	For Each objItem in colItems
	    Txt.WriteLine "Name: " & objItem.Name
	    Txt.WriteLine "Build number: " & objItem.BuildNumber
	    Txt.WriteLine "Caption: " & objItem.Caption
	    Txt.WriteLine "Characteristics: " & objItem.Characteristics
	    Txt.WriteLine "CSD version: " & objItem.CSDVersion
	    Txt.WriteLine "Dedicated: " & objItem.Dedicated
	    Txt.WriteLine "Description: " & objItem.Description
	    Txt.WriteLine "Flags: " & objItem.Flags
	    Txt.WriteLine "Identifying descriptions: " & objItem.IdentifyingDescriptions
	    Txt.WriteLine "Initial load info: " & objItem.InitialLoadInfo
	    Txt.WriteLine "Installation date: " & objItem.InstallDate
	    Txt.WriteLine "Last load info: " & objItem.LastLoadInfo
	    Txt.WriteLine "Major version: " & objItem.MajorVersion
	    Txt.WriteLine "Minor version: " & objItem.MinorVersion
	    Txt.WriteLine "Name format: " & objItem.NameFormat
	    Txt.WriteLine "Node highest version: " & objItem.NodeHighestVersion
	    Txt.WriteLine "Node lowest version: " & objItem.NodeLowestVersion
	    Txt.WriteLine "Other identifying info: " & objItem.OtherIdentifyingInfo
	    Txt.WriteLine "Power state: " & objItem.PowerState
	    Txt.WriteLine "Primary owner contact: " & objItem.PrimaryOwnerContact
	    Txt.WriteLine "Primary owner name: " & objItem.PrimaryOwnerName
	    Txt.WriteLine "Reset capability: " & objItem.ResetCapability
	    Txt.WriteLine "Roles: " & objItem.Roles
	    Txt.WriteLine "State: " & objItem.State
	    Txt.WriteLine "Status: " & objItem.Status
	    Txt.WriteLine ""
	Next
Else 
	WScript.Echo "This computer is not a node of a cluster."
End If 

Txt.Close


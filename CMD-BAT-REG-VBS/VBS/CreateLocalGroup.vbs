'=====================================================================
'
' This Script will create a local Group on the Server where it is running
'
' Script created by Holger Habermehl. October 23, 2012
'=====================================================================

	Dim DefaultGroup
	Dim GroupName

	DefaultGroup = "LocalUserGroup"   			' Set default value.

	' Display A message box where you can insert the group name, leave it blank to create the group defined under DefaultGroup
	GroupName = InputBox("What's the Name of the new local User Group? The default is: " & DefaultGroup, "Create Group", GroupName)
	IF GroupName = "" then GroupName = DefaultGroup

	' Confirmation window for the new local group
    MyVar = MsgBox ("Do you want to create the group " & GroupName & " to your local system?", 36, "Run Script")
    If MyVar<> 6 then WScript.Quit

	' Call for function to create the local group
	CreateLocalGroup GroupName
	
	' Tell the user that the script was done
    WScript.Echo "Script Completed!"

	' Needed functions
	Function CreateLocalGroup(strLGroupName)
         strComputer = GetComputerName()
         Set colAccounts = GetObject("WinNT://" & strComputer & "")
         Set objUser = colAccounts.Create("group", strLGroupName)
         objUser.SetInfo
	End Function

	Function GetComputerName()
		Set WSHNetwork = WScript.CreateObject("WScript.Network")
		strComputerName  = wshNetwork.ComputerName
		GetComputerName = strComputerName
	End Function
	
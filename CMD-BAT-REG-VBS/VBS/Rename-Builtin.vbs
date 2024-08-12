strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colAccounts = objWMIService.ExecQuery _
    ("Select * From Win32_UserAccount Where LocalAccount = True And (Name = 'Guest' or Name ='Administrator')")

For Each objAccount in colAccounts
	If objAccount.name="Guest" Then
		objaccount.Rename "Visitor"
	End If 
	
	If objAccount.name="Administrator" Then
		objAccount.Rename "rotatrtsinimda"
	End If
Next
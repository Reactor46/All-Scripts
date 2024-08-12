'=*=*=*=*=*=*=*=*=*=*=*=
' Coded By Assaf Miron 
' Date : 07/02/08
'=*=*=*=*=*=*=*=*=*=*=*=
'*================================================*
'* This Script Changes the Administrator Password to a Complex Password.
'* The Password Change is Made to the Default SID of the Administrator User.
'* The Administrator Password is Saved to a Log File for Later Review 
'* (You can Remove the Logging by Setting the LOG_Cahnges to False)
'* You Can Change the Script To Run on a Local Computer or on a Remote Computer by Setting the strComputer Const
'*================================================*
'================================================
' Consts
'================================================
Const LogName = "ChangeAdminPass"
Const ForAppending = 8
Const strComputer = "." 'Local Computer
Const LOG_Changes = True

'================================================
' Dims and Variables
'================================================
Dim objFSO
Dim objFile

'================================================
' Sets
'================================================
If LOG_Changes Then
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(LogName & ".txt", ForAppending, True)
End If

'================================================
' Functions
'================================================
Function ScramblePass(Pass)
' This Function Scrambels a Text
'	Function Recivies : String Password
'	Function Returns : Scrambled String Password	
	Randomize
	
	intPassLen = Len(Pass)
	For i = 1 to intPassLen
	' Create a Random Value Between 1 to 10
		MyValue = Int((10 * Rnd) + 1)  
	' Read the Given Password Character by Character to the tmpPass
		tmpPass = Mid(Pass,i,1)
	' Add the Modifiend Character to the tPass
		tPass = Chr(Asc(tmpPass) + MyValue) & tPass
	Next
	ScramblePass = tPass
End Function

'================================================
' Subs
'================================================
Sub ChangeAdminPass(Pass)
' This Sub Checks the Administrators User Name by his SID 
' and Changes his Password
'	Sub Recieves : a String Password
	' Get a Scrambled Password
	Pass = ScramblePass(Pass)
	' Connnect to the Target Computer's WMI and Search all the Local Users
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")	
	Set colAccounts = objWMIService.ExecQuery _
	    ("Select * From Win32_UserAccount Where Domain = '" & strComputer & "'")
	
	For Each objAccount in colAccounts
	' Search for the User Accout that his SID Begins with S-1-5- and Ends with -500
	    If Left (objAccount.SID, 6) = "S-1-5-" and Right(objAccount.SID, 4) = "-500" Then
	    	If LOG_Changes Then _
	        	objFile.WriteLine objAccount.Name
        ' Set the Scrambled Password to the User
        objAccount.SetPassword Pass
	    End If
	Next
	
	If LOG_Changes Then _
		objFile.WriteLine "Administrators Password Will Be : " & Pass
	' Set the User Information and the New Password
	Err = objUser.SetInfo
	If Err <> 0 Then
		If LOG_Changes Then _
			objFile.WriteLine "Could not change Administrator's Password"
	Else
		If LOG_Changes Then _
			objFile.WriteLine "Administrator's Password Has been Changed"
	End If
End Sub


ChangeAdminPass "P@ssw0rd1"
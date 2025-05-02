' ------------------------------------------------------------------------------------
' The sample scripts are not supported under any Microsoft standard support 
' program or service. The sample scripts are provided AS IS without warranty 
' of any kind. Microsoft further disclaims all implied warranties including, 
' without limitation, any implied warranties of merchantability or of fitness for 
' a particular purpose. The entire risk arising out of the use or performance of 
' the sample scripts and documentation remains with you. In no event shall 
' Microsoft, its authors, or anyone else involved in the creation, production, or 
' delivery of the scripts be liable for any damages whatsoever (including, 
' without limitation, damages for loss of business profits, business interruption, 
' loss of business information, or other pecuniary loss) arising out of the use 
' of or inability to use the sample scripts or documentation, even if Microsoft 
' has been advised of the possibility of such damages
' ------------------------------------------------------------------------------------

Option Explicit
On Error Resume Next

Const HKEY_CURRENT_USER   = &H80000001
Const strWin8Key = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer"
Const strComputer = "."
Dim objReg : Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
Dim objWMIService
Dim colItems
Dim colItem
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

' Check if registry key exists.
Function RegistryKeyExists(strRoot, strKey, strSubkey)
	RegistryKeyExists = False
	Dim aSubkeys, s
	objReg.EnumKey strRoot, strKey, aSubkeys 
	If Not IsNull(aSubkeys) Then 
		For Each s In aSubkeys 
			If LCase(s)=LCase(strSubkey) Then
				RegistryKeyExists = True 
				Exit Function 
			End If 
		Next 
	End If
End Function

' Change background color of Windows 8 Start screen with user input.
Function SetOSCWin8StartScreenBackColor(colorIndex)
	If RegistryKeyExists(HKEY_CURRENT_USER,strWin8Key,"Accent") Then
		' If registry key exists then set value with user input.
		objReg.SetDWORDValue HKEY_CURRENT_USER,strWin8Key & "\" & "Accent","ColorSet_Version3",colorIndex
	Else
		' If registry key does not exist, create registry key and set value with user input.
		objReg.CreateKey HKEY_CURRENT_USER, strWin8Key & "\" & "Accent"
		objReg.SetDWORDValue HKEY_CURRENT_USER,strWin8Key & "\" & "Accent","ColorSet_Version3",colorIndex
	End If
	
	' Require user confirm if log off now.
	Dim logoffNow
	logoffNow = MsgBox(REQUIRELOGOFF,4,"")
	If logoffNow = 6 Then
		For Each colItem In colItems
			' Force current user log off.
			colItem.Win32Shutdown(4)
		Next
	Else
		WScript.Echo REQUIRELOGOFFMANUALLY
	End If
End Function

' Set background color of Windows 8 Start screen to default(Option 8).
Function SetOSCWin8StartScreenBackColorDefault
	SetOSCWin8StartScreenBackColor 8
End Function

' Import all strings used by this script.
Sub ReadAllStrings
	Dim fso : Set fso = WScript.CreateObject("Scripting.FileSystemObject")
	Dim stringPath
	stringPath = fso.GetParentFolderName(Wscript.ScriptFullName) & "\en-US\Strings.vbs"
	ExecuteGlobal fso.OpenTextFile(stringPath, 1).ReadAll
	Set fso = Nothing
End Sub

' Check if current Operating System is Windows 8.
Function IsWin8()
	IsWin8 = False
	For Each colItem In colItems
		If Not IsNull(colItem.Version) Then
			If Len(colItem.Version) > 3 Then
				If Left(colItem.Version,3) = "6.2" Then
					IsWin8 = True
				End If
			End If
		End If
	Next
End Function

'Entry point of script execution.
Sub Main
	
	Call ReadAllStrings()
	
	If IsWin8() Then 
		' Verify user input.
		Dim input
		input = InputBox(INPUTDESCRIPTION,INPUTTITLE,8)
		If input <> "" Then
			Dim colorOption
			colorOption = input
			If LCase(colorOption) = Lcase("default") Then
				'Set background color of Windows 8 Start screen to default
				Call SetOSCWin8StartScreenBackColorDefault
			Else
				If IsNumeric(colorOption) Then
					If colorOption >= 0 And colorOption <= 24 Then
						'Set background color of Windows 8 Start screen with user input.
						SetOSCWin8StartScreenBackColor colorOption
					Else
						' Range of user input is 0~24.
						WScript.Echo PARAMETER_RANGE_ERROR
					End If
				Else
					' Integer user input is required.
					WScript.Echo PARAMETER_NUMERIC_REQUIRED
				End If
			End If
		End If
	Else
		'Please run this script on Windows 8.
		WScript.Echo REQUIREWIN8
	End If
	
End Sub

Call Main()
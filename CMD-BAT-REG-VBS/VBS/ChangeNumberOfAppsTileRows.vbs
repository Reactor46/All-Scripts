'#--------------------------------------------------------------------------------- 
'#The sample scripts are not supported under any Microsoft standard support 
'#program or service. The sample scripts are provided AS IS without warranty  
'#of any kind. Microsoft further disclaims all implied warranties including,  
'#without limitation, any implied warranties of merchantability or of fitness for 
'#a particular purpose. The entire risk arising out of the use or performance of  
'#the sample scripts and documentation remains with you. In no event shall 
'#Microsoft, its authors, or anyone else involved in the creation, production, or 
'#delivery of the scripts be liable for any damages whatsoever (including, 
'#without limitation, damages for loss of business profits, business interruption, 
'#loss of business information, or other pecuniary loss) arising out of the use 
'#of or inability to use the sample scripts or documentation, even if Microsoft 
'#has been advised of the possibility of such damages 
'#--------------------------------------------------------------------------------- 

Option Explicit

Dim strAnswer,RowNumber,objShell

Set objShell = CreateObject("Wscript.Shell")

strAnswer = InputBox("Please input the number you want to change the number of tile rows:")
If strAnswer = "" Then
	WScript.Quit
Else
	RowNumber = CInt(strAnswer)
End If 

'Define a key registry path
Dim regMaximumRowsPath,regRowsCounPath,MaximumRowValue

regMaximumRowsPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\ImmersiveShell\Grid\Layout_MaximumAvailableHeightCells"
regRowsCounPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\ImmersiveShell\Grid\Layout_MaximumRowCount"

MaximumRowValue = objShell.RegRead(regMaximumRowsPath)

'Setting the registry key
If RowNumber <= MaximumRowValue And RowNumber >=1 Then
	objShell.RegWrite regRowsCounPath,RowNumber,"REG_DWORD"
	WScript.Echo "Setting the number of tile row successfully."

'Call function
	Choice
Else
	WScript.Echo "The number of tile rows are not less than minimum value ""1"" and are not greater than maximum value """ & MaximumRowValue & """, please re-enter the value."
End If

'Prompt message
Sub Choice
Dim result

	result = MsgBox ("It will take effect after log off, do you want to log off right now?", vbYesNo, "Log off computer")
	
	Select Case result
	Case vbYes
		objShell.Run("logoff")
	Case vbNo
		Wscript.Quit
	End Select
End Sub	
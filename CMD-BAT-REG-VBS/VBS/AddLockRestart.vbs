'---------------------------------------------------------------------------------
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
' has been advised of the possibility of such damages.
'---------------------------------------------------------------------------------
Option Explicit 
'On Error Resume Next 

Sub  main()
	Const HKEY_CLASSES_ROOT = &H80000000
	Dim count 
	'get the count of arguments
   	count = WScript.Arguments.Count
   	Select Case count 
   		Case 1 
   			Dim Item
   			Item = UCase(WScript.Arguments.Item(0))
   			'verify the value of item
   			Select Case Item 
   				'if value equals 'add',add the items into context menu.
   				Case "ADD"
   					'add 'lock computer' to context menu
					If KeyExists("HKEY_CLASSES_ROOT\DesktopBackground\Shell\lock Computer\") = False   Then 
						CreateKeys "Lock Computer","shell32.dll,-325","Rundll32 User32.dll,LockWorkStation"
					End If 
					'add 'restart computer' 
					If KeyExists("HKEY_CLASSES_ROOT\DesktopBackground\Shell\Restart Computer\") = False   Then 
						CreateKeys "Restart Computer","shell32.dll,-221","shutdown.exe -r -t 00 -f"
					End If
					'add 'sleep computer'
					If KeyExists("HKEY_CLASSES_ROOT\DesktopBackground\Shell\Sleep Computer\") = False   Then 
						CreateKeys "Sleep Computer","shell32.dll,-331","rundll32.exe powrprof.dll,SetSuspendState 0,1,0"  
					End If
					'add 'shutdown computer'
					If KeyExists("HKEY_CLASSES_ROOT\DesktopBackground\Shell\Shutdown Computer\") = False   Then 
						CreateKeys "Shutdown Computer","shell32.dll,-329","shutdown.exe -s -t 00 -f" 
					End If		
					WScript.Echo "Add 'Lock,Restart,Sleep,ShutDown Computer' to context menu successfully"
					
   				Case "REMOVE"
   					'Remove the items 
   					DeleteSubkeys 	HKEY_CLASSES_ROOT,"DesktopBackground\Shell\LOCK COMPUTER"
   					DeleteSubkeys 	HKEY_CLASSES_ROOT,"DesktopBackground\Shell\RESTART COMPUTER"
   					DeleteSubkeys	HKEY_CLASSES_ROOT,"DesktopBackground\Shell\SLEEP COMPUTER"
   					DeleteSubkeys	HKEY_CLASSES_ROOT,"DesktopBackground\Shell\SHUTDOWN COMPUTER"	
   					WScript.Echo "Remove 'Lock,Restart,Sleep,ShutDown Computer' from context menu successfully."			
   				Case Else
   					WScript.Echo "Invalid argument, please try again."
   					
   			End Select 
   		
   		Case Else 
  			WScript.Echo "Invalid argument, please try again."
   	End Select 		
End Sub 


'################################################
' This script is to create registry key and its subkey
'################################################
Function CreateKeys(Name,IconValue,CommandValue)
	Const HKEY_CLASSES_ROOT = &H80000000
	Dim strComputer, strValueName, objRegistry,strKeyPath, path
	path = "HKEY_CLASSES_ROOT\DesktopBackground\Shell\" & Name & "\"
	strKeyPath = "DesktopBackground\Shell\" & Name
	strComputer = "."
	strValueName = ""
	Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	If KeyExists(Path)  Then 
		CreateKeys  = False 
	Else 
		objRegistry.CreateKey HKEY_CLASSES_ROOT, strKeyPath   'Create a registry key
		objRegistry.SetStringValue HKEY_CLASSES_ROOT, strKeyPath, "Icon", IconValue
		objRegistry.SetStringValue HKEY_CLASSES_ROOT, strKeyPath, "Position", "Bottom"
		objRegistry.CreateKey HKEY_CLASSES_ROOT, strKeyPath & "\Command"
		objRegistry.SetStringValue HKEY_CLASSES_ROOT, strKeyPath & "\Command", strValueName, CommandValue
		If KeyExists(Path) Then  'Verify if the key is created
			CreateKeys  = True 
		Else 
			CreateKeys  = False 
		End If 
	End If	
End Function


'################################################
' This script is to verify if the registry key exists
'################################################
Function KeyExists(Path)
	On Error Resume Next 
	Dim objshell,Flag,value
	Set objShell = CreateObject("WScript.Shell")
	value = objShell.RegRead(Path) 
	Flag = False 
	If Err.Number = 0 Then 	
	 	Flag = True 
	End If
	Keyexists = Flag
End Function 

'################################################
'This function is to delete registry key.
'################################################
Sub DeleteSubkeys(HKEY_CLASSES_ROOT,strKeyPath) 
	Dim strSubkey,arrSubkeys,strComputer,objRegistry
	strComputer = "."
	Set objRegistry = GetObject("winmgmts:\\" & _
    strComputer & "\root\default:StdRegProv")
    objRegistry.EnumKey HKEY_CLASSES_ROOT, strKeyPath, arrSubkeys 
    If IsArray(arrSubkeys) Then 
        For Each strSubkey In arrSubkeys 
            DeleteSubkeys HKEY_CLASSES_ROOT, strKeyPath & "\" & strSubkey 
        Next 
    End If 
    objRegistry.DeleteKey HKEY_CLASSES_ROOT, strKeyPath 
End Sub


Call main 
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

On Error Resume Next

Dim regKeyPath,value,filesizeValue
value = Inputbox("ID ShortcutKey" & vbCrLf &_                                                                            
                 "-- -----------" & vbCrLf &_                                                                          
                  "1 Win+A" & vbCrLf &_                                                                                  
                  "2 Win+B" & vbCrLf &_                                                                                  
                  "3 Win+C" & vbCrLf &_                                                                                  
                  "4 Win+D" & vbCrLf &_                                                                                  
                  "5 Win+E" & vbCrLf &_                                                                                  
                  "6 Win+F" & vbCrLf &_                                                                                  
                  "7 Win+G" & vbCrLf &_                                                                                  
                  "8 Win+H" & vbCrLf &_                                                                                  
                  "9 Win+I" & vbCrLf &_                                                                                  
                  "10 Win+J" & vbCrLf &_                                                                                  
                  "11 Win+K" & vbCrLf &_                                                                                  
                  "12 Win+L" & vbCrLf &_                                                                                  
                  "13 Win+M" & vbCrLf &_                                                                                  
                  "14 Win+N" & vbCrLf &_                                                                                  
                  "15 Win+O" & vbCrLf &_                                                                                  
                  "16 Win+P" & vbCrLf &_                                                                                  
                  "17 Win+Q" & vbCrLf &_                                                                                  
                  "18 Win+R" & vbCrLf &_                                                                                  
                  "19 Win+S" & vbCrLf &_                                                                                  
                  "20 Win+T" & vbCrLf &_                                                                                  
                  "21 Win+U" & vbCrLf &_                                                                                  
                  "22 Win+V" & vbCrLf &_                                                                                  
                  "23 Win+W" & vbCrLf &_                                                                                  
                  "24 Win+X" & vbCrLf &_                                                                                  
                  "25 Win+Y" & vbCrLf &_                                                                                  
                  "26 Win+Z" & vbCrLf &_
"Which shortcut key do you want to set for OneNote? Input the corresponding ID number")

Select Case value
      case "1" dValue = 65
      		   Main
      case "2" dValue = 66
      		   Main
      case "3" dValue = 67
               Main
      case "4" dValue = 68
               Main
      case "5" dValue = 69
               Main
      case "6" dValue = 70
               Main
      case "7" dValue = 71
               Main
      case "8" dValue = 72
               Main
      case "9" dValue = 73
               Main
      case "10" dValue = 74
               Main
      case "11" dValue = 75
               Main
      case "12" dValue = 76
               Main
      case "13" dValue = 77
               Main
      case "14" dValue = 78
               Main
      case "15" dValue = 79
               Main
      case "16" dValue = 80
               Main
      case "17" dValue = 81
               Main
      case "18" dValue = 82
               Main
      case "19" dValue = 83
               Main
      case "20" dValue = 84
               Main
      case "21" dValue = 85
               Main
      case "22" dValue = 86
               Main
      case "23" dValue = 87
               Main
      case "24" dValue = 88
               Main
      case "25" dValue = 89
               Main
      case "26" dValue = 90
      Case Else MsgBox "You input the wrong option, please enter again"
End Select

Sub Main
Dim objRegistry,strValueName,dwValue,strComputer
	Const HKEY_CURRENT_USER = &H80000001
	strComputer = "."
	
	Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	strKeyPath = "Software\Microsoft\Office\15.0\OneNote\Options\Other"
	objRegistry.CreateKey HKEY_CURRENT_USER, strKeyPath
	
	strValueName = "ScreenClippingShortcutKey"
	dwValue = dValue
	
	objRegistry.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, strValueName, dwValue
	WScript.Echo "Set the value of properties of ScreenClippingShortcutKey successfully."
	
	'Call function
	Choice
End Sub

'Prompt message
Sub Choice
	Dim result,objShell
	
	Set objShell = WScript.CreateObject("WScript.Shell")
	result = MsgBox ("It will take effect after log off, do you want to log off right now?", vbYesNo, "Log off computer")
	
	Select Case result
	Case vbYes
		objShell.Run("logoff")
	Case vbNo
		Wscript.Quit
	End Select
End Sub	

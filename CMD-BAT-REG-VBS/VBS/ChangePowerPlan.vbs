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

Dim strComputer
Dim colItems,colItem
Dim objWMIService
Dim regKeyPath,value,filesizeValue
Dim objWshShell

value = Inputbox("ID PowerOption" & vbCrLf &_                                                                            
                 "-- -----------" & vbCrLf &_                                                                          
                 "1  Balanced" & vbCrLf &_                                                                                  
                 "2  High performance" & vbCrLf &_                                                                                  
                 "3  Powersaver" & vbCrLf &_                                                                                  
"Which power option do you want to set ? Input the corresponding ID number")

Set objWshShell = WScript.CreateObject("WScript.Shell")

If IsEmpty(value) Then
	WScript.Quit
Else
	Select Case value
	      Case "1" objWshShell.Run "powercfg -s 381b4222-f694-41f0-9685-ff5bb260df2e"
	      	CheckCurrentPowerPlan
	      		   
	      Case "2" objWshShell.Run "powercfg -s 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
	      	CheckCurrentPowerPlan
	      		   
	      Case "3" objWshShell.Run "powercfg -s a1841308-3541-4fab-bc81-f71556f20b4a"
	      	CheckCurrentPowerPlan
	               
	      Case Else MsgBox "You input the wrong option, please enter again"
	End Select
	
	Wscript.Sleep 2000
End If

Function CheckCurrentPowerPlan
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer &"\root\cimv2\power")
	Set colItems = objWMIService.ExecQuery("Select * From Win32_PowerPlan where isActive='true'")
	
	For Each colItem in colItems
	    Wscript.Echo "Now, the current power plan option is " & colItem.ElementName
	Next
End Function
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
Dim objShell,objExecObject

Set objShell = WScript.CreateObject("WScript.Shell")

Set objExecObject = objShell.Exec("Dism /online /Get-FeatureInfo /FeatureName:Internet-Explorer-Optional-amd64")

Dim Flag,strText,returnValue
Dim resultRemoveIE

Flag = 0
Do While Not objExecObject.StdOut.AtEndOfStream
    strText = objExecObject.StdOut.ReadLine()
    returnValue = InStr(strText,"Enabled")
	Flag = Flag + returnValue
Loop


If Flag > 0 Then
	resultRemoveIE = MsgBox("Are you sure you want to remove Internet Explorer. This process might take a long time, so please be patient.",vbYesNo+vbQuestion,"Remove Internet Explorer")
	If resultRemoveIE = 6 Then
		Set objExecObject = objShell.Exec("Dism /online /Disable-Feature /FeatureName:Internet-Explorer-Optional-amd64")	
		While objExecObject.Status = 0
			WScript.Sleep 1
		Wend
		Call ChoicePromt
	End If
Else
	resultRemoveIE = MsgBox("Internet Explorer is already removed, do you want to install Internet Explorer? This process might take a long time, so please be patient.",vbYesNo+vbQuestion,"Install Internet Explorer")
	If resultRemoveIE = 6 Then
		Set objExecObject = objShell.Exec("Dism /online /Enable-Feature /FeatureName:Internet-Explorer-Optional-amd64")
		While objExecObject.Status = 0
			WScript.Sleep 1
		Wend
		Call ChoicePromt
	End If
End If

Sub ChoicePromt
	Dim resultReboot
	
	'reboot the computer
	resultReboot = MsgBox("You may need to reboot of windows for the change to take effect. Do you want to reboot the computer right now?",vbYesNo+vbQuestion,"Reboot Computer")
	If resultReboot = 6 Then
		objShell.Exec("shutdown -r -t 0")
	End If
End Sub

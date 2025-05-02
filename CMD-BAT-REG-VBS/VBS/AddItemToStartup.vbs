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
On Error Resume Next

Dim UserConfigDestination,objShell,WshShell,envAppDATA

Set objShell = CreateObject("Shell.Application")  
Set WshShell = CreateObject("WScript.Shell")

envAPPDATA = WshShell.expandEnvironmentStrings("%APPDATA%")

userConfigDestination = envAPPDATA & "\Microsoft\Windows\Start Menu\Programs\Startup\"

Dim ArgCount,File,FSO 
ArgCount = WScript.Arguments.Count
 
Select Case ArgCount 
        Case 1  
        	File = WScript.Arguments(0)
        		Set FSO = CreateObject("Scripting.FileSystemObject")
        			FSO.CopyFile File,UserConfigDestination
        			If Err.Number = 0 Then
        				WScript.Echo "Add file successfully."
        			Else
        				WScript.Echo "Please drag a file to this script." 
        			End If
        Case  Else 
                WScript.Echo "Please drag a file to this script." 
End Select 

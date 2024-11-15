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

Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")

Dim strPath,objExecObject
WScript.StdOut.Write("Please type the path of custom refresh image you want to save: ")
strPath = WScript.StdIn.ReadLine()

'call recimg.exe to create custom image
Set objExecObject = objShell.Exec("recimg /createimage " & strPath)
Do 
	WScript.StdOut.WriteLine(objExecObject.StdOut.ReadLine())
Loop While Not objExecObject.StdOut.AtEndOfStream

WScript.StdOut.WriteLine(objExecObject.StdOut.ReadAll)

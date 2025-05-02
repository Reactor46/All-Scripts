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

Dim objShell
Dim objShellWindows
Dim objFSO
Dim objItem
Dim objFile
Dim objFileName
Dim LocationURL
Dim LocationName
Dim objExportFile
Dim currentDirectory

'Create an instance 
Set objShell = CreateObject("Shell.Application")
'Then use the Windows method to get all window object
Set objShellWindows = objShell.Windows()

'Create an instance of the Scripting.FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
Const ForWriting = 2
'Get current folder location
currentDirectory = Left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))

If objShellWindows.Count = 0 Then
	WScript.Echo "Can not find any opend tabs in Internet Explorer."
Else
	Set objExportFile = objFSO.CreateTextFile(currentDirectory & "IEurls.csv", ForWriting, True)
	objExportFile.WriteLine "URL" & "," & "Title"
	
	For Each objItem In objShellWindows
		FileFullName = objItem.FullName
		objFile = objFSO.GetFile(objItem.FullName)
		objFileName = objFSO.GetFileName(objFile)
		
		If LCase(objFileName) = "iexplore.exe" Then
			LocationURL = objItem.LocationURL
			LocationName = objItem.LocationName
			objExportFile.WriteLine LocationURL & "," & LocationName
		End If
	Next
	
	WScript.Echo "Successfully generated IEurls.csv on " & currentDirectory
End If
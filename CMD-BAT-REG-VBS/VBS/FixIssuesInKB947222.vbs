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

Dim strFiles,ShowSubFolders
strGroup = "administrators"
SE_DACL_PRESENT = &h4
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.Shell")

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True
objRegEx.IgnoreCase = True
objRegEx.Pattern = "Access is denied"

'call the InputBox function
strAnswer = InputBox("Please input the folder location you want to search:","Folder Location")
If strAnswer = "" Then
    Wscript.Quit
Else
    objStartFolder = strAnswer
End If

'Checks if a folder exists
If objFSO.FolderExists(objStartFolder) Then 
    For Each Subfolder in objFSO.GetFolder(objStartFolder).SubFolders
    	'Checks if a file exists
        strFiles = Subfolder+"\desktop.ini"
        If objFSO.FileExists(strFiles) Then
        	Set objExea = objShell.Exec("icacls " & strFiles)
			msg=objExea.StdErr.ReadLine()
	
			If objRegEx.Test(msg) Then
				WScript.Echo Subfolder & " - Not Modified!"
			Else
				Set objExeb = objShell.Exec("icacls " & strFiles & " /deny " & strGroup & ":R")
					If objExeb.ExitCode = 0 Then 
						WScript.Echo Subfolder & " - Success!"
					Else
						errMsg = objExeb.StdErr.ReadAll()
						WScript.Echo errMsg
					End If
			End If
		Else
			WScript.Echo strFiles & " does not exist."
		End If
    Next
Else
	WScript.Echo objStartFolder & " does not exist."
End If 
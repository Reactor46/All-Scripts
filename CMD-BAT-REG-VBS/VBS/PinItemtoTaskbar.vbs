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

Set objShell = CreateObject("Shell.Application")  

Dim ArgCount
Dim File,FileName
ArgCount = WScript.Arguments.Count
 
Select Case ArgCount 
	Case 1  
    	File = WScript.Arguments.Item(0)
		FileName = File
		Do While InStr(FileName,"\") > 0
			FileName = Right(FileName,Len(FileName)-InStr(FileName,"\"))
		Loop 
    
    	'Pin file to taskbar
		Dim objFolder,objFolderItem,colVerbs,objVerb
		
		Set objFolder = objShell.Namespace(Left(File,Len(File)-Len(FileName)))
		Set objFolderItem = objFolder.ParseName(FileName) 
		Set colVerbs = objFolderItem.Verbs
		
		'Verify the file can be pinned to taskbar
		Dim Flag
		Flag=0
			
		For Each objVerb in colVerbs	
			If Replace(objVerb.name, "&", "") = "Pin to Taskbar" Then 
				objVerb.DoIt
				Flag = 1
			End If
		Next
			
		If Flag = 1 Then 
			msgbox "Pin '"& FileName &"' file to taskbar successfully."
		Else 
			msgbox "Failed to pin '"& FileName &"' file to taskbar successfully."
		End If 		        	
    Case  Else 
        WScript.Echo "Please drag a file to this script." 
End Select 






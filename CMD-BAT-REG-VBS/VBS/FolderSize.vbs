'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 2009
'
' NAME:Folder size 
'
' AUTHOR: Mohammed Alyafae , 
' DATE  : 9/22/2011
'
' COMMENT: this script uses very intelligent and fast way to calculate folder size
' it uses the output of  windows command "dir /s /a >c:\size.txt"  and extracts the size value
' from the end of the output file 
' it is very fast and efficient way to calulate folder size
'==========================================================================
Option Explicit
On Error Resume Next
Dim objShell
Dim objcmd
Dim strcmd
Dim FolderPath

FolderPath=InputBox("Enter Folder Path","Check Folder Size","C:\")
strcmd="cmd /c " & "dir " & FolderPath & " /s /a" & "> c:\size.txt" 'command is dir /s /a > c:\size.txt

Set objShell=CreateObject("WScript.Shell")
objShell.Run strcmd,0,True  ' I use here Run instead of exec cos Run has the Third param to wait untill command finished

ExtractSize


Sub ExtractSize
Dim strLine 'to store next line from the file
Dim iFirst
Dim iLast
Dim iLength
Dim objFSO
Dim objFile


Set objFSO=CreateObject("Scripting.FileSystemObject")
Set objFile=objFSO.OpenTextFile("c:\size.txt",1)

Do While objFile.AtEndOfStream <> True

strLine=objFile.ReadLine

If InStr(strLine,"Total Files Listed")<> 0 Then  
	strLine=objFile.ReadLine
	WScript.Echo strLine
	iFirst=InStr(strLine,")")
	iLast=InStr(strLine,"bytes")
	iFirst=iFirst+1
	iLast=iLast-1
	iLength=iLast-iFirst	
	WScript.Echo round(Mid(strLine,iFirst,iLength)/1048576) & " MB"
	
End If

Loop

End Sub
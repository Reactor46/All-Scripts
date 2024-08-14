'#######################################################################
'#       Author: Vikas Sukhija 
'#       Date: 04/28/2011
'#       Description: Launch Multiple Robocopy session & log only
'#       differences 
'#######################################################################

Option Explicit
 
Dim objShell, objFSO
Dim strOptions, strLog, strSource, strDestination, strFolder
Dim txtFile, strFileName, strPath
Dim objBaseFolder, objFolder
Dim Strlogfind, strsource1, strDestination1, strFolder1
Set objShell = CreateObject("WScript.Shell")      
Set objFSO = CreateObject("Scripting.FileSystemObject") 
     
Set txtFile =       objFSO.OpenTextFile("C:\scripts\robocopy\readfile.txt", 1)   ' read from text file –change accordingly
 
do until txtFile.AtEndOfStream

strFolder      = txtFile.ReadLine   'read folder names

strFolder1 = chr(34) & strFolder & chr(34)  'add quotes to avoid space errors
 
 
  strSource      = "\\source-fs\c$\Scripts" & "\" & strFolder      'Source (Change accordingly)

  strsource1 =  chr(34) & strSource & chr(34)  'add quotes
 
  strDestination = "\\destination-fs\c$\backup" & "\" & strFolder      'Destination (Change accordingly)

  strDestination1 = chr(34) & strDestination & chr(34) 'add quotes
 
  objShell.run "cmd /k C:\scripts\robocopy\rcpy.bat "& " " & strSource1 & " " & strDestination1 & " " & strFolder1 &""  'batch file command

  WScript.Sleep 60000  ' launch session every 60 secs

loop

txtFile.Close

'###########################################################################
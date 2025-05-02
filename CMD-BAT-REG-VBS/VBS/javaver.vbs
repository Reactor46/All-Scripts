'-----------------------------------------------------------------------------------------------------------
' Script: javaver.vbs
' Author: Jeff Mason aka TNJMAN aka bitdoctor
' Date:   09/18/2013
' Purpose: To check Java.exe version on a list of remote computers
'
' Put Java version in "jver" variable
' Put File location underneath "program files" that you are searching in "floc1" variable
' Put File name (Java.exe, in this case) in the "fnm1" variable
' Put Report/Log name ("java-version-report.txt" in this case) in "rnm1" variable
' Put Report/Log location in "rloc1" variable
' NOTE: If report file does not exist, it will be created initially, then appeneded to
'
' Assumptions: 
' 1) You have (or create) a c:\scripts folder to hold the script and the output
' 2) You have the needed privs to run this against the remote workstation
' 3) You MODIFY the variables to meet your needs & environment
' NOTE: This script can be used/modified to check versions of MS Office, Acrobat Flash, etc.
'       Since it is "comma-separated, you can easily import into Excel or OpenOffice Calc, etc.
'
' To run this script, use either #1 or #2 below
'
' 1) For a single computer, at command prompt: cscript javaver.vbs remote-computer-name //nologo
'    Example: cscript c:\scripts\javaver.vbs remotepc1 //nologo
' 2) To run against multiple computers, 
'     a) Create c:\scripts\java-report.bat with multiple computer names
'        Example of 'java-report.bat'
'         rem -- java-report.bat [calling file] --
'         cscript c:\scripts\javaver.vbs TJONES //nologo
'         cscript c:\scripts\javaver.vbs PSMITH //nologo
'         cscript c:\scripts\javaver.vbs DLITTLE //nologo
'         rem -- where TJONES, PSMITH & DLITTLE are examples of remote computer names
'     b) Execute/run the calling bat file; example:
'         c:\scripts\java-report.bat
'         
' 3) When the run (#1 or #2 above) has finished, review the report/log:
'    From command prompt, type "notepad c:\scripts\java-version-report.txt" to check the resulting versions
'
'-----------------------------------------------------------------------------------------------------------
'
Dim wshShell
Dim strComputer, windir
Dim fso, rptlog
Dim targetfile, fvar, fversion, fname, fdtmod
Dim fnm1, jver, floc1, rpt
'
'-------------------------------------
'MODIFY THESE VARIABLES TO YOUR NEEDS
'-------------------------------------
'
jver = "jre7"
floc1 = "Java\" & jver & "\bin\"
fnm1 = "Java.exe"
rnm1 = "java-version-report.txt"
rloc1 = "c:\scripts\"
rpt = rloc1 & rnm1

Const ForAppending = 8

strComputer = Wscript.Arguments.Item(0)

' --> for debugging: Wscript.Echo strComputer

Set wshShell = WScript.CreateObject ("WScript.Shell")

windir = wshShell.ExpandEnvironmentStrings ("%WINDIR%")

Set fso = CreateObject ("Scripting.FileSystemObject")

set rptlog = fso.OpenTextFile (rpt, ForAppending, TRUE)

Set objFSO = CreateObject("Scripting.FileSystemObject")

'64 or 32 bit
If objFSO.FolderExists("\\" & strComputer & "\c$\windows\syswow64") Then
    targetfile = "\\" & strComputer & "\c$\program files (x86)\" & floc1 & fnm1 & ""
Else
    targetfile = "\\" & strComputer & "\c$\program files\" & floc1 & fnm1 & ""
End If

' --> for debugging: msgbox(targetfile)

On Error Resume Next

Set fvar = fso.GetFile (targetfile)
' for debugging: msgbox "Error: " & err.Number
If err.Number = 0 Then
  fversion = fso.GetFileVersion (targetfile)
  fname = fvar.name
  fdtmod = fvar.DateLastModified
Else
  err.Clear
  fversion = "Unknown version"
  fname = "No Java version " & jver
  fdtmod = "Unknown modified date"
End If

On Error Goto 0

' Write a line and then continue on to check for JRE v6
rptlog.write strComputer & ", " & fname & ", " & fdtmod & ", " & fversion & VbCrLf

jver = "jre6"
floc1 = "Java\" & jver & "\bin\"
fnm1 = "Java.exe"
rnm1 = "java-version-report.txt"
rloc1 = "c:\scripts\"
rpt = rloc1 & rnm1

'64 or 32 bit
If objFSO.FolderExists("\\" & strComputer & "\c$\windows\syswow64") Then
    targetfile = "\\" & strComputer & "\c$\program files (x86)\" & floc1 & fnm1 & ""
Else
    targetfile = "\\" & strComputer & "\c$\program files\" & floc1 & fnm1 & ""
End If

' --> for debugging: msgbox(targetfile)

On Error Resume Next

Set fvar = fso.GetFile (targetfile)
' for debugging: msgbox "Error: " & err.Number
If err.Number = 0 Then
  fversion = fso.GetFileVersion (targetfile)
  fname = fvar.name
  fdtmod = fvar.DateLastModified
Else
  err.Clear
  fversion = "Unknown version"
  fname = "No Java version " & jver
  fdtmod = "Unknown modified date"
End If

On Error Goto 0

' Write a line and close the rpt/log file
rptlog.write strComputer & ", " & fname & ", " & fdtmod & ", " & fversion & VbCrLf
rptlog.close : set rptlog = nothing


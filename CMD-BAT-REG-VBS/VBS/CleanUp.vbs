'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 2009
'
' NAME:Windows CleanUp 
'
' AUTHOR: Mohammed Alyafae , 
' DATE  : 10/1/2011
'
' COMMENT: This script used "cleanmgr /sagerun:1" command to perform cleanup for  the following windows temporary files
' Active Setup Temp Folders
' Content Indexer Cleaner
' Downloaded Program Files
' GameNewsFiles
' GameStatisticsFiles
' GameUpdateFiles
' Internet Cache Files
' Memory Dump Files
' Microsoft Office Temp Files
' Offline Pages Files
' Old ChkDsk Files
' Previous Installations
' Recycle Bin
' Service Pack Cleanup
' Setup Log Files
' System error memory dump files
' System error minidump files
' Temporary Files
' Temporary Setup Files
' Temporary Sync Files
' Thumbnail Cache
' Upgrade Discarded Files
' Windows Error Reporting Archive Files
' Windows Error Reporting Queue Files
' Windows Error Reporting System Archive Files
' Windows Error Reporting System Queue Files
' Windows Upgrade Log Files
'
' First you have to set registry  key stateFlags001 to DWORD value 2
' in this registery path HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches
'then run this command cleanmgr /sagerun:1
'
'==========================================================================
Option Explicit
On Error Resume Next

SetRegKeys
DoCleanup


Sub DoCleanup()
Dim WshShell
Set WshShell=CreateObject("WScript.Shell")
WshShell.Run "C:\WINDOWS\SYSTEM32\cleanmgr /sagerun:1"
End Sub

Sub SetRegKeys
Dim strKeyPath
Dim strComputer
Dim objReg
Dim arrSubKeys
Dim SubKey
Dim strValueName
Const HKLM=&H80000002


strKeyPath="SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
strComputer="."
strValueName="StateFlags0001"

Set objReg=GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
objReg.Enumkey HKLM ,strKeyPath,arrSubKeys

For Each SubKey In arrSubKeys

objReg.SetDWORDValue HKLM,strKeyPath & "\" & SubKey,strValueName,2

Next

End Sub

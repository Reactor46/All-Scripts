' February 2013
' by Romano Jerez
'
' OK to use and distribute if author information is kept
' ===========================================================

Option Explicit

Dim strPCList, objFSO, objWshNet, strLocalPC, objPCFile, strCurrentLine
Dim arrayName, strOldName, strNewName, intExitCode, objComputer 
Dim objWMIService, Return, bFound, strLocalLogFile, objLocalLogFile
Dim strUserAccount, strUserPassword, strRemoteLogFileSuccess
Dim strRemoteLogFileFail, strRemoteLogFileNoRename, objRemoteLogFileSuccess
Dim objRemoteLogFileFail, objRemoteLogFileNoRename, bSameName
Dim strRemoteLogFileSuccessShare, strRemoteLogFileFailShare
Dim strRemoteLogFileNoRenameShare


Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8


' *** Configure file with computer list ***

  strPCList = "\\<serverName>\<shareName>\renameList.txt"

' *****************************************


' *** Configure local log File and remote log locations ***
'
  strLocalLogFile = "c:\windows\temp\pcRename.log"
  strRemoteLogFileSuccessShare = "\\<serverName>\<shareName>\Success"
  strRemoteLogFileFailShare = "\\<serverName>\<shareName>\Fail"
  strRemoteLogFileNoRenameShare = "\\<serverName>\<shareName>\NoRenameAttempt"
'
' ********************************

' *** Configure Account with permissions to rename computer in AD ***
'
  strUserAccount = "<domainName>\<domainAccount>"
  strUserPassword = "<password>"
'
' ********************************

Set objWshNet = CreateObject("WScript.Network")
Set objFSO = CreateObject("Scripting.FileSystemObject")

strLocalPC = objWshNet.ComputerName

strRemoteLogFileSuccess = strRemoteLogFileSuccessShare & "\" & strLocalPC & ".txt"
strRemoteLogFileFail = strRemoteLogFileFailShare & "\" & strLocalPC & ".txt"
strRemoteLogFileNoRename = strRemoteLogFileNoRenameShare & "\" & strLocalPC & ".txt"

initializeLogging



Set objPCFile = objFSO.OpenTextFile(strPCList, ForReading)

intExitCode = 99

bSameName = vbFalse
bFound = vbFalse 

Do Until (objPCFile.AtEndOfStream)
  strCurrentLine = objPCFile.Readline

  If Len(strCurrentLine) > 0 Then
    arrayName = Split(strCurrentLine, ",", -1, 1)
    strOldName = Trim(arrayName(0))
    strNewName = Trim(arrayName(1))

    If strLocalPC = strOldName Then
      bFound = vbTrue
      If (Len(strNewName)) > 15 Then
        WScript.Echo "Not renaming as new name has more than 15 characters: " & strNewName
        writeToLog("Not renaming as new name has more than 15 characters: " & strNewName)
        LogNoRename("Not renaming as new name has more than 15 characters: " & strNewName)       
      ElseIf strOldName = strNewName Then
        bSameName = vbTrue
        WScript.Echo "Not renaming as old and New name are the same for " & strLocalPC
        writeToLog("Not renaming as old and New name are the same for " & strLocalPC)
        LogNoRename("Not renaming as old and New name are the same for " & strLocalPC)
      Else
        WScript.Echo "Old name: " & strOldName & ".  Changing to " & strNewName
        writeToLog("Old name: " & strOldName & ".  Changing to " & strNewName)
        intExitCode = Rename(strNewName)
      End If
    End If
  End If
Loop



If Not bFound Then
  WScript.Echo strLocalPC & " not found in list as old computer. No renaming attempted."
  writeToLog(strLocalPC & " not found in list as old computer.  No renaming attempted.")
  LogNoRename(strLocalPC & " not found in list as old computer.  No renaming attempted.")
End If


If intExitCode = 99 Then  ' didn't run Rename function
  intExitCode = 0         ' but return that the script ran successfully
End If

cleanUp


WScript.Quit(intExitCode)



' *********************************************************************

Sub cleanUp

  WScript.Echo "Exit code: " & intExitCode
  writeToLoG("Exit code: " & intExitCode)
  objPCFile.Close
  objLocalLogFile.Close

End Sub

' *****************

Sub initializeLogging

  If objFSO.FileExists(strLocalLogFile) Then
    Set objLocalLogFile = objFSO.OpenTextFile(strLocalLogFile, ForAppending)
  Else
    Set objLocalLogFile = objFSO.CreateTextFile(strLocalLogFile)
    objLocalLogFile.Close
    Set objLocalLogFile = objFSO.OpenTextFile(strLocalLogFile, ForWriting)
  End If

  objLocalLogFile.WriteLine(Now())

End Sub

' *****************

Sub LogNoRename(strMessage)

  If objFSO.FileExists(strRemoteLogFileNoRename) Then
    Set objRemoteLogFileNoRename = objFSO.OpenTextFile(strRemoteLogFileNoRename, ForAppending)
  Else
    Set objRemoteLogFileNoRename = objFSO.CreateTextFile(strRemoteLogFileNoRename)
    objRemoteLogFileNoRename.Close
    Set objRemoteLogFileNoRename = objFSO.OpenTextFile(strRemoteLogFileNoRename, ForWriting)
  End If

  objRemoteLogFileNoRename.WriteLine(Now())
  objRemoteLogFileNoRename.WriteLine(strMessage)
  objRemoteLogFileNoRename.Close

End Sub

' ******************

Sub LogRenameFailed(strMessage)

  If objFSO.FileExists(strRemoteLogFileFail) Then
    WScript.Echo "File exists"
    Set objRemoteLogFileFail = objFSO.OpenTextFile(strRemoteLogFileFail, ForAppending)
  Else
    Set objRemoteLogFileFail = objFSO.CreateTextFile(strRemoteLogFileFail)
    objRemoteLogFileFail.Close
    Set objRemoteLogFileFail = objFSO.OpenTextFile(strRemoteLogFileFail, ForWriting)
  End If

  objRemoteLogFileFail.WriteLine(Now())
  objRemoteLogFileFail.WriteLine(strMessage)
  objRemoteLogFileFail.Close

End Sub

' ******************

Sub LogRenameSucceded(strMessage)

  If objFSO.FileExists(strRemoteLogFileSuccess) Then
    Set objRemoteLogFileSuccess = objFSO.OpenTextFile(strRemoteLogFileSuccess, ForAppending)
  Else
    Set objRemoteLogFileSuccess = objFSO.CreateTextFile(strRemoteLogFileSuccess)
    objRemoteLogFileSuccess.Close
    Set objRemoteLogFileSuccess = objFSO.OpenTextFile(strRemoteLogFileSuccess, ForWriting)
  End If

  objRemoteLogFileSuccess.WriteLine(Now())
  objRemoteLogFileSuccess.WriteLine(strMessage)
  objRemoteLogFileSuccess.Close

End Sub

' ****************************************

Function Rename(strName)
  Set objWMIService = GetObject("Winmgmts:root\cimv2") 
 
  For Each objComputer in objWMIService.InstancesOf("Win32_ComputerSystem")
    Return = objComputer.Rename(strName,strUserPassword,strUserAccount)
    
    If Return <> 0 Then
      WScript.Echo "Rename to " & strName & " failed. Error = " & Err.Number
      writeToLog("Rename to " & strName & " failed. Error = " & Err.Number)
      LogRenameFailed("Rename to " & strName & " failed. Error = " & Err.Number & vbCrLf & "Exit code: " & Return)
    Else
      WScript.Echo "Rename to " & strName & " succeeded. Restart is needed."
      writeToLog("Rename to " & strName & " succeeded. Restart is needed.")
      LogRenameSucceded("Rename to " & strName & " succeeded. Restart is needed.")
    End If
  
  Next

  Rename = Return
End Function


' ***************************************

Sub writeToLog(strMessage)

  objLocalLogFile.WriteLine(strMessage)

End Sub

' ***************************************
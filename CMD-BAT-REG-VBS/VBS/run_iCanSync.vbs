'##################################################################################################
'## 
'##      Copyright (c) 2010 TCPconsulting, LLC.
'##                  All Rights Reserved.
'## 
'## 
'## 
'## DISCLAIMER:
'## 
'##      Disclaimer of Warranty
'##
'##      This code and information is provided "AS IS" without warranty of any kind, 
'## either expressed or implied, including but not limited to the implied warranties 
'## of merchantability and/or fitness for a  particular purpose. Good data processing procedure
'## dictates that any program/code be thoroughly tested with non-critical data before relying on it. 
'## The user must assume the entire risk of using the program. The developer does not retain 
'## any liability for data loss, loss of profits or any other kind of damage caused through
'## the use of this product.
'##
'##
'##   Version: 5.4.17 [19 Oct 2011 10:23]
'##
'##
'##################################################################################################



Option Explicit

' #####
' ## vars
' #####

Const cAppDir = "iCanSyncEZ"
Const cfnMkdirCmdFile = "ics_exportMakePath"
Const cfnPackExclude = "ics_packExclude"
Const cfnUnpackIgnore = "ics_unpackIgnore" ' ***
Const cConfigFileExt = ".config"
Const cfnAppOutFile = "out"
Const cLinksRoot = "_ics_links"
Const cMaxRunTime = 40
Const cOldStateDir = "States"
Const cStateExt = ".state"
Const cfnCurrSuff = "_curr"
Const cfnImportedSuff = "_import"
Dim cCurrStateExt: cCurrStateExt = cStateExt & cfnCurrSuff
Dim cAppDataPath: cAppDataPath = "%APPDATA%\"& cAppDir
Dim cAppLogFile: cAppLogFile = cAppDir &".log"
Dim cAppSetupFile: cAppSetupFile = cAppDir &"Setup.xml"
Dim cSetupApp: cSetupApp = cAppDir &"Setup.html"
Dim cRestoreConfigFile: cRestoreConfigFile = cAppDir &"Restore.xml"
Dim cLinkOwnerChangeFile: cLinkOwnerChangeFile = cAppDir &"LinkOwnerChanges.xml"
Dim cfnInProcess: cfnInProcess = cAppDir &"_inProcess.txt"
Dim cJobCmdFilePath: cJobCmdFilePath = "%ALLUSERSPROFILE%\"& cAppDir
Dim cIcsTmpMark: cIcsTmpMark = "_icsTmp"

Class CIcsSyncType
   Dim importOnly
   Dim backupOnly
   Dim exportOnly
   Dim synchronize

   Private Sub Class_Initialize
      importOnly = "importOnly"
      backupOnly = "backupOnly"
      exportOnly = "exportOnly"
      synchronize = "exportImport"
   End Sub
End Class

Class CIcsLinkType
   Dim directory
   Dim file

   Private Sub Class_Initialize
      directory = "directory"
      file = "file"
   End Sub
End Class

Class CAltTransportPath
   Dim web
   Dim ftp

   Private Sub Class_Initialize
      web = "webPath:"
      ftp = "remotePath:"
   End Sub
End Class


Dim parm_bLinkActive, parm_linkType, parm_linkName, parm_syncType, parm_syncPass
Dim parm_ftp, parm_user, parm_pass
Dim parm_smtpServer, parm_smtpUser, parm_smtpPass, parm_emailFrom, parm_emailTo
Dim parm_dirDestRoot
Dim parm_bImportOnlyForceRefresh, parm_hImportJobToRunBefore, parm_hImportJobToRunAfter, parm_bImportAltTransportUseFlag, parm_importAltTransportPath
Dim parm_hExportJobToRunBefore, parm_hExportJobToRunAfter, parm_exportMaxFileSize
Dim parm_bPackFileImportExport, parm_bExportImportFirstTime
Dim parm_bBackUpRunImport, parm_bBackUpRunExport, parm_backUpDir, parm_backUpCount
Dim parm_archType
Dim parm_oFilesToProcess
Dim parm_aExportExcludePath, parm_aExportExcludeFile
Dim parm_aImportExcludePath, parm_aImportExcludeFile
Dim parm_aExportIgnoreChangedFile
Dim parm_bNoTopLevelFolders, parm_bNoTopLevelFiles ' !!! export only right now

Dim gl_scriptWorkDir, gl_packExcludeLoc, gl_mkdirCmdFileLoc, gl_appLoc
Dim gl_unpackIgnoreLoc ' ***
Dim gl_stateFile, gl_stateFileCurr, gl_stateFileCurrLoc
Dim gl_linkImportFile, gl_linkImportFileLoc, gl_linkExportFile, gl_linkExportFileLoc, gl_linkRemoteRoot, gl_linkDirName, gl_toPackExt
Dim gl_syncRootDirExpand
Dim gl_fShowReport
Dim gl_appSetupFilePath, gl_scriptErrBuff
Dim gl_aFullyUsedDirs, gl_aFullyUsedFiles

Dim gl_restoreLinkName, gl_restoreFileName, gl_restoreBackUpFileName
Dim gl_jobCmdFilePath

gl_tcpc_debug = True


' #####
' ## setup
' #####

On Error Resume Next

const HKCU = &H80000001
const HKLM = &H80000002

Dim oReg: Set oReg = GetObject ("winmgmts:\\.\root\default:StdRegProv")
Dim sKeyPathAuthorizedApps: sKeyPathAuthorizedApps = "SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\AuthorizedApplications\List"
Dim sValueNameAuthorizedApps: sValueNameAuthorizedApps = "%WinDir%\\system32\\FTP.exe"
Dim sValueAuthorizedApps: sValueAuthorizedApps ="%WinDir%\\system32\\FTP.exe:*:Enabled:File Transfer Program"
oReg.SetStringValue HKLM, sKeyPathAuthorizedApps, sValueNameAuthorizedApps, sValueAuthorizedApps
Set oReg = Nothing


Dim gl_oXml: Set gl_oXml = WScript.CreateObject("Microsoft.XMLDOM")
Dim gl_oFso: Set gl_oFso = WScript.CreateObject ("Scripting.FileSystemObject")
Dim gl_oShell: Set gl_oShell = WScript.CreateObject ("WScript.Shell")
Dim gl_oStringManip: Set gl_oStringManip = New CStringManip
Dim gl_oSupport: Set gl_oSupport = New CSupport
Dim gl_icsSyncType: Set gl_icsSyncType = New CIcsSyncType
Dim gl_icsLinkType: Set gl_icsLinkType = New CIcsLinkType
Dim gl_icsAltTransportPath: Set gl_icsAltTransportPath = New CAltTransportPath
Dim gl_oFileStruct: Set gl_oFileStruct = New CFileStructure
Dim gl_fLinkOwnerChanged: gl_fLinkOwnerChanged = False

Set parm_oFilesToProcess = CreateObject("Scripting.Dictionary")
Set parm_hImportJobToRunBefore = CreateObject("Scripting.Dictionary")
Set parm_hImportJobToRunAfter = CreateObject("Scripting.Dictionary")
Set parm_hExportJobToRunBefore = CreateObject("Scripting.Dictionary")
Set parm_hExportJobToRunAfter = CreateObject("Scripting.Dictionary")


Dim gl_scriptLoc: gl_scriptLoc = WScript.ScriptFullName
Dim gl_scriptName: gl_scriptName = WScript.ScriptName
Dim gl_scriptPath: gl_scriptPath = gl_oFso.GetFile (gl_scriptLoc).ParentFolder
Dim gl_scriptOutLoc: gl_scriptOutLoc = gl_scriptPath &"\"& cfnAppOutFile

gl_appSetupFilePath = muExpandEnvironmentStrings (cAppDataPath, gl_scriptErrBuff)


' ## check & create job cmd path
gl_jobCmdFilePath = gl_oShell.ExpandEnvironmentStrings (cJobCmdFilePath)
Call gl_oSupport.makeDir (gl_oFso, gl_jobCmdFilePath)
If Err Then Call TheEnd (19020, Err.Number & ": " & Err.Description)


If Not gl_oFso.FileExists (gl_appSetupFilePath &"\"& cAppSetupFile) Then
   Dim startMenuPath, setupAppMenuPath

   startMenuPath = gl_oShell.SpecialFolders ("StartMenu")
   setupAppMenuPath = startMenuPath &"\Programs\"& cAppDir

   Call gl_oSupport.makeDir (gl_oFso, gl_appSetupFilePath &"\"& cOldStateDir)
   If Err Then Call TheEnd (19019, Err.Number & ": " & Err.Description)

   If gl_scriptErrBuff = "" Then
      Call gl_oSupport.makeDir (gl_oFso, setupAppMenuPath)
      If Err Then Call TheEnd (19016, Err.Number & ": " & Err.Description)

      Dim oShortcutUrl: Set oShortcutUrl = gl_oShell.CreateShortcut (setupAppMenuPath &"\"& cAppDir &" - Control Panel.url")
      oShortcutUrl.TargetPath = gl_appSetupFilePath &"\"& cSetupApp
      oShortcutUrl.Save

      Set oShortcutUrl = gl_oShell.CreateShortcut (setupAppMenuPath &"\View Log.url")
      oShortcutUrl.TargetPath = gl_appSetupFilePath &"\"& cAppLogFile
      oShortcutUrl.Save
      Set oShortcutUrl = Nothing

      gl_oFso.CopyFile gl_scriptPath &"\"& cSetupApp, gl_appSetupFilePath &"\"& cSetupApp, True
      gl_oFso.CopyFile gl_scriptPath &"\"& cAppSetupFile, gl_appSetupFilePath &"\"& cAppSetupFile, True

      Dim xmlRoot: Set xmlRoot = gl_oXml.createElement ("iCanSyncAppLoc")
      Dim xmlNode: Set xmlNode = gl_oXml.createElement ("AppLoc")
      
      Dim parentInfo: Set parentInfo = gl_oSupport.getScriptParentProcessInfo ()
      xmlNode.Text = parentInfo.ExecutablePath

      gl_oXml.documentElement = xmlRoot
      xmlRoot.appendChild (xmlNode)

      gl_oXml.save (gl_oFso.BuildPath (gl_appSetupFilePath, cAppDir &"AppLoc.xml"))

      gl_oShell.Run "cmd /c """& gl_appSetupFilePath &"\"& cSetupApp &"""", 0, False
   End If

   Call cleanUp ()

   If gl_scriptErrBuff = "" Then
      TheEnd 0, "The End" ' we need to run setup first
   Else
      TheEnd 19017, "Cannot select correct user to run job."
   End If
End If


' ## check GUI
If Not gl_oFso.FileExists (gl_appSetupFilePath &"\"& cSetupApp) Then
   gl_oFso.CopyFile gl_scriptPath &"\"& cSetupApp, gl_appSetupFilePath &"\"& cSetupApp, True
End If


gl_restoreLinkName = ""
gl_restoreFileName = ""
gl_restoreBackUpFileName = ""

' ## check for Restore file
If gl_oFso.FileExists (gl_appSetupFilePath &"\"& cRestoreConfigFile) Then
   Dim oXmlRestore: Set oXmlRestore = WScript.CreateObject("Microsoft.XMLDOM")

   oXmlRestore.async="false"
   oXmlRestore.Load (gl_appSetupFilePath &"\"& cRestoreConfigFile)

   If oXmlRestore.parseError.errorCode Then
      Call TheEnd (19011, "Parse Error: "& _ 
               oXmlRestore.parseError.reason & vbcrlf & _ 
               " Line = "    & oXmlRestore.parseError.line & vbcrlf & _ 
               " linePos = " & oXmlRestore.parseError.linePos & vbcrlf & _ 
               " srcText = " & oXmlRestore.parseError.srcText & vbcrlf & _ 
               " ErrorCode = " & oXmlRestore.parseError.ErrorCode & vbcrlf)
   End If

   gl_restoreLinkName = oXmlRestore.selectSingleNode("//iCanSyncRestore/iCanSyncLink/linkName").Text
   gl_restoreFileName = oXmlRestore.selectSingleNode("//iCanSyncRestore/iCanSyncLink/fileName").Text
   gl_restoreBackUpFileName = oXmlRestore.selectSingleNode("//iCanSyncRestore/iCanSyncLink/backUpFileName").Text

   Set oXmlRestore = Nothing
   Call gl_oFso.DeleteFile (gl_appSetupFilePath &"\"& cRestoreConfigFile, True)
End If


Call log2file ("APPLICATION STARTED (5.4.17 [19 Oct 2011 10:23])", "I")

gl_scriptWorkDir = gl_scriptPath &"\"& cLinksRoot
gl_packExcludeLoc = gl_scriptWorkDir &"\"& cfnPackExclude
gl_unpackIgnoreLoc = gl_scriptWorkDir &"\"& cfnUnpackIgnore ' ***
gl_mkdirCmdFileLoc = gl_scriptWorkDir &"\"& cfnMkdirCmdFile
gl_appLoc = gl_scriptPath &"\"& cAppDir
gl_scriptErrBuff = ""
gl_fShowReport = False
gl_toPackExt = ".wz"

Call gl_oSupport.makeDir (gl_oFso, gl_scriptWorkDir)
If Err Then Call TheEnd (19012, Err.Number & ": " & Err.Description)


' ## to prevent from using same directory for different instances
If gl_oFso.FileExists (gl_scriptPath &"\"& cfnInProcess) Then
   Dim oInProcessFile: Set oInProcessFile = gl_oFso.GetFile (gl_scriptPath &"\"& cfnInProcess)
   Dim timeDiff: timeDiff = Abs (DateDiff ("n", Now, oInProcessFile.DateLastModified)) ' h - Hour; n - Minute
   Set oInProcessFile = Nothing

   If timeDiff < cMaxRunTime Then Call TheEnd (19015, "Another instance of the program is running for "& timeDiff & " min.")
End If

Call gl_oFso.CreateTextFile (gl_scriptPath &"\"& cfnInProcess, True, 0)
' ## /to prevent from using same directory for different instances

Call log2file ("Cleaning up ...", "I")
Call cleanUp ()
Call log2file ("Cleaning is completed", "I")

gl_oXml.async="false"
gl_oXml.Load (gl_appSetupFilePath &"\"& cAppSetupFile)

If gl_oXml.parseError.errorCode Then
   Call TheEnd (19001, "Parse Error: "& _ 
            gl_oXml.parseError.reason & vbcrlf & _ 
            " Line = "    & gl_oXml.parseError.line & vbcrlf & _ 
            " linePos = " & gl_oXml.parseError.linePos & vbcrlf & _ 
            " srcText = " & gl_oXml.parseError.srcText & vbcrlf & _ 
            " ErrorCode = " & gl_oXml.parseError.ErrorCode & vbcrlf)
End If

Dim oNodes: Set oNodes = gl_oXml.selectNodes ("//iCanSyncSetup/iCanSyncLink")
Dim oNode, retIsRunNeeded


' #####
' ## main
' #####

For Each oNode In oNodes
   Call setDefaultSetup ()
   Call getSetup ()
'   MsgBox "Main for: "& parm_linkName &", linkType: "& parm_linkType

   If Err Then 
      reportProblem ()
   Else
      Call log2file ("sync started -- link: "& parm_linkName &", linkType: "& parm_linkType &", syncType: "& parm_syncType, "I")

      Dim errFlag

      gl_stateFile = parm_linkName & cStateExt
      gl_stateFileCurr = parm_linkName & cCurrStateExt
      gl_stateFileCurrLoc = gl_scriptWorkDir &"\"& gl_stateFileCurr
      gl_linkImportFile = parm_linkName & cfnImportedSuff & gl_toPackExt
      gl_linkImportFileLoc = gl_scriptWorkDir &"\"& gl_linkImportFile
      gl_linkExportFile = parm_linkName & gl_toPackExt
      gl_linkExportFileLoc = gl_scriptWorkDir &"\"& gl_linkExportFile

      ' if gl_linkDirName = "" then this is not ics sync -- importAltTransport || file AltTransport export
      If InStr (parm_importAltTransportPath, gl_icsAltTransportPath.ftp) = 1 Then
         gl_linkDirName = ""
         gl_linkRemoteRoot = Replace (parm_importAltTransportPath, gl_icsAltTransportPath.ftp, "")
      ElseIf parm_syncType = gl_icsSyncType.importOnly And parm_bImportAltTransportUseFlag And InStr (parm_importAltTransportPath, gl_icsAltTransportPath.web) = 1 Then
         gl_linkRemoteRoot = Replace (parm_importAltTransportPath, gl_icsAltTransportPath.web, "")

         If InStr (parm_importAltTransportPath, cLinksRoot) - 1 + Len (cLinksRoot) = Len (parm_importAltTransportPath) Then
            gl_linkDirName = parm_linkName &"_"& parm_linkType
         Else
            gl_linkDirName = ""
         End If
      Else
         gl_linkDirName = parm_linkName &"_"& parm_linkType
         gl_linkRemoteRoot = cLinksRoot
      End If

      retIsRunNeeded = isRunNeeded ()
      errFlag = reportProblem ()

      If retIsRunNeeded > 0 Then
         Call log2file ("sync required -- link: "& parm_linkName, "I")

         If parm_syncType = gl_icsSyncType.backupOnly Then
            Call runBackUp (True)
            errFlag = reportProblem ()

            Call runJob ("export", "RunAfter")
         ElseIf parm_syncType = gl_icsSyncType.importOnly Then
            Call runImport (retIsRunNeeded)
            errFlag = reportProblem ()
         ElseIf parm_syncType = gl_icsSyncType.exportOnly Then
            Call runExport (retIsRunNeeded)
            errFlag = reportProblem ()
         ElseIf parm_syncType = gl_icsSyncType.synchronize Then
            Call log2file (gl_icsSyncType.synchronize &" -- currently is not supported", "W")

            ' Call runExportImport (retIsRunNeeded)
            ' errFlag = reportProblem ()
         End If

         If errFlag = 0 And parm_syncType <> gl_icsSyncType.backupOnly And gl_linkDirName <> "" And parm_bImportAltTransportUseFlag = False And gl_restoreFileName = "" Then
            updateConfigFile ()
         End If

         Call log2file ("sync FINISHED -- link: "& parm_linkName &", type: "& parm_linkType, "I")
      Else
         If parm_syncType = gl_icsSyncType.exportOnly Or parm_syncType = gl_icsSyncType.backupOnly Then Call runJob ("export", "RunAfter")

         Call log2file ("sync FINISHED -- everything is up today for link: "& parm_linkName &", type: "& parm_linkType, "I")
      End If
   End If
Next

TheEnd 0, "The End"


' #####
' ## TheEnd (ByVal in_te_iErrCode, ByVal in_te_sMsg)
Sub TheEnd (ByVal in_te_iErrCode, ByVal in_te_sMsg)
   Dim te_logLoc, te_newOutFile
   '
   ' Clean up
   On Error Resume Next

   If in_te_iErrCode > 0 Then Call log2file (in_te_sMsg, "E")
   Call log2file ("APPLICATION FINISHED", "I")

   te_logLoc = gl_appSetupFilePath &"\"& cAppLogFile

   Set gl_oXml = Nothing
   Set oNodes = Nothing
   Set oNode = Nothing
   Set oRegEx = Nothing

   If in_te_iErrCode <> 19015 Then
      Call gl_oFso.DeleteFolder (gl_scriptWorkDir, True)
      Call gl_oFso.DeleteFile (gl_scriptPath &"\"& cfnInProcess, True)
      Call gl_oFso.DeleteFile (gl_scriptOutLoc, True)
   End If

   If in_te_iErrCode > 0 Or gl_fShowReport = True Then gl_oShell.Run "notepad "& te_logLoc, 1, False

   Set gl_oFso = Nothing
   Set gl_oShell = Nothing
   Set gl_oStringManip = Nothing
   Set gl_oSupport = Nothing
   Set gl_oFileStruct = Nothing
   Set gl_icsSyncType = Nothing
   Set gl_icsLinkType = Nothing
   Set gl_icsAltTransportPath = Nothing

   WScript.Quit (in_te_iErrCode)
End Sub


' #####
' ## log2file (ByVal in_lf_msg, ByVal in_lf_msgType)
Function log2file (ByVal in_lf_msg, ByVal in_lf_msgType)
   Dim lf_currDate, lf_currMonth, lf_currDay, lf_ts
   Dim lf_oStream, lf_logLoc

   On Error Resume Next

   lf_logLoc = gl_appSetupFilePath &"\"& cAppLogFile

   in_lf_msg = gl_oStringManip.replaceRegExp (in_lf_msg, "\n", " ")

   log2file = 0

   lf_currDate = Now
   lf_currMonth = DatePart ("m" , lf_currDate)
   lf_currDay = DatePart ("d" , lf_currDate)

   If lf_currMonth < 10 Then lf_currMonth = "0"& lf_currMonth

   If lf_currDay < 10 Then lf_currDay = "0"& lf_currDay
   
   lf_ts = DatePart ("yyyy" , lf_currDate) & lf_currMonth & lf_currDay &" "& FormatDateTime (Time,4)

   If in_lf_msgType = "E" Then
      lf_ts = lf_ts &" -E- "
      gl_fShowReport = True
   ElseIf in_lf_msgType = "W" Then
      lf_ts = lf_ts &" -W- "
   Else 
      lf_ts = lf_ts &"     "
   End If
   
   Set lf_oStream = gl_oFso.OpenTextFile (lf_logLoc, 8, True, -1)

   lf_oStream.Write lf_ts & in_lf_msg
   If in_lf_msgType = "E" Then lf_oStream.Write " "& gl_scriptErrBuff
   If in_lf_msgType = "E" And gl_tcpc_debug = True Then lf_oStream.Write " "& tcpc_debugGetStack()
   lf_oStream.Write vbcrlf
   lf_oStream.Close

   Set lf_oStream = Nothing

   If Err Then
      log2file = "ERROR " & Err.Number & ": " & Err.Description &" @log2file"
      Err.Clear
   End If

   On Error Goto 0
End Function


' #####
' ## runBackUp (ByVal in_rbu_bPack)
Sub runBackUp (ByVal in_rbu_bPack)
   Dim fn: fn = "runBackUp"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   If in_rbu_bPack Then Call pack ()

   moveToBackUp ()

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Sub


' #####
' ## runImport (ByVal in_ri_isRunNeeded)
Sub runImport (ByVal in_ri_isRunNeeded)
   Dim ri_i

   Dim fn: fn = "runImport"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   If parm_bBackUpRunImport Then Call runBackUp (True)

   If in_ri_isRunNeeded = 1 Then 
      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
      Exit Sub
   End If

   If parm_linkType = gl_icsLinkType.file Then
      For Each ri_i In parm_oFilesToProcess.Keys
         If parm_oFilesToProcess(ri_i).importRequest = True Then
            If parm_bPackFileImportExport Then
               If getFile (ri_i & gl_toPackExt, ri_i & gl_toPackExt & cfnImportedSuff, gl_linkDirName) = True Then
                  Call ics_unpack (ri_i & gl_toPackExt & cfnImportedSuff, ri_i & gl_toPackExt)
               Else
                  Call log2file (" # runImport: file is not accessible ["& ri_i &"]", "W")
               End If
            ElseIf gl_linkDirName = "" Then ' importAltTransport
               Call gl_oSupport.makeDir (gl_oFso, gl_syncRootDirExpand)
               Call gl_oFso.CopyFile (gl_scriptWorkDir &"\"& ri_i & cfnImportedSuff, gl_syncRootDirExpand &"\"& ri_i)
            Else
               If getFile (ri_i, ri_i & cfnImportedSuff, gl_linkDirName) = True Then
                  Call gl_oSupport.makeDir (gl_oFso, gl_syncRootDirExpand)
                  Call gl_oFso.CopyFile (gl_scriptWorkDir &"\"& ri_i & cfnImportedSuff, gl_syncRootDirExpand &"\"& ri_i)

                  If gl_restoreLinkName = "" And Not (Not parm_bPackFileImportExport And InStr (ri_i, gl_toPackExt) + 2 = Len (ri_i)) Then
                     Dim ri_oState: Set ri_oState = WScript.CreateObject("Microsoft.XMLDOM")
                     ri_oState.async = "false"
                     ri_oState.Load (gl_scriptWorkDir &"\"& ri_i & cStateExt)

                     If ri_oState.parseError.errorCode Then
                        Call log2file (" # runImport: failed to open file -- "& ri_status &". Line: "& oXmlRestore.parseError.line &", srcText: "& oXmlRestore.parseError.srcText, "E")
                     Else
                        Call gl_oSupport.touchFile (gl_syncRootDirExpand, ri_i, ri_oState.selectSingleNode("//root/files/file/modifiedDate").Text)
                     End If

                     Set ri_oState = Nothing
                  End If
               Else
                  Call log2file (" # runImport: file is not accessible ["& ri_i &"]", "W")
               End If
            End If
         End If
      Next
   Else
      If getFile (gl_linkExportFile, gl_linkImportFile, gl_linkDirName) = True Then
         Call ics_unpack (gl_linkImportFile, gl_linkExportFile)
      Else
         Call log2file (" # runImport: file is not accessible ["& gl_linkExportFile &"]", "W")
      End If
   End If

   Call runJob ("import", "RunAfter")

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Sub


' #####
' ## runExport (ByVal in_re_isRunNeeded)
Sub runExport (ByVal in_re_isRunNeeded)
   Dim re_i, re_oFile
   Dim re_exportCount: re_exportCount = 0

   Dim fn: fn = "runExport"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   Call pack()
   
   Dim re_localFileLoc, re_remoteFileLoc, re_currFileSize

   If parm_linkType = gl_icsLinkType.file Then
      For Each re_i In parm_oFilesToProcess.Keys
         If parm_bPackFileImportExport Then
            re_localFileLoc = gl_scriptWorkDir &"\"& re_i & gl_toPackExt
            re_remoteFileLoc = re_i & gl_toPackExt
         Else
            re_localFileLoc = gl_syncRootDirExpand &"\"& re_i
            re_remoteFileLoc = re_i
         End If

         Set re_oFile = gl_oFso.GetFile (re_localFileLoc) 
         re_currFileSize = Int (re_oFile.Size / 1024)

         If re_currFileSize < Int (parm_exportMaxFileSize) Or Int (parm_exportMaxFileSize) = 0 Then
            If putFile (re_localFileLoc, re_remoteFileLoc, gl_linkDirName) Then re_exportCount = re_exportCount + 1
            If gl_linkDirName <> "" Then Call putFile (gl_scriptWorkDir &"\"& re_i & cCurrStateExt, re_i & cStateExt, gl_linkDirName)
         Else
            Call log2file ("Size of "& re_localFileLoc &" is "& re_currFileSize &"KB. Max size: "& parm_exportMaxFileSize &"KB", "E")
            gl_fShowReport = True
         End If
      Next
   Else 
      If putFile (gl_linkExportFileLoc, gl_linkExportFile, gl_linkDirName) Then re_exportCount = re_exportCount + 1
      If gl_linkDirName <> "" Then Call putFile (gl_stateFileCurr, gl_stateFile, gl_linkDirName)
   End If

   If parm_bBackUpRunExport Then Call runBackUp (False)

   Call runJob ("export", "RunAfter")

   If gl_fLinkOwnerChanged = True And re_exportCount > 0 Then ' gl_fLinkOwnerChanged will be set to True only if gl_linkDirName <> ""
      Dim re_oXmlLinkOwnerChange: Set re_oXmlLinkOwnerChange = WScript.CreateObject("Microsoft.XMLDOM")
      Dim re_oNodeLinkOwnerChange

      re_oXmlLinkOwnerChange.async = "false"
      re_oXmlLinkOwnerChange.Load (gl_oFso.BuildPath (gl_appSetupFilePath, cLinkOwnerChangeFile))

      If Not re_oXmlLinkOwnerChange.parseError.errorCode Then
         Set re_oNodeLinkOwnerChange = re_oXmlLinkOwnerChange.selectSingleNode("//linkOwnerChanges/"& UCase (gl_linkDirName))
         re_oNodeLinkOwnerChange.parentNode.removeChild (re_oNodeLinkOwnerChange)

         re_oXmlLinkOwnerChange.save (gl_oFso.BuildPath (gl_appSetupFilePath, cLinkOwnerChangeFile))
      End If

      Call log2file ("Link's ownership has been taking.", "W")

      Set re_oNodeLinkOwnerChange = Nothing
      Set re_oXmlLinkOwnerChange = Nothing
   ElseIf gl_fLinkOwnerChanged = True And re_exportCount = 0 Then
      Call log2file ("Link's ownership has NOT been taking from previous owner because export failed.", "W")
   End If

'   Dim oNet: Set oNet = CreateObject("WScript.NetWork")
'   Dim msgText: msgText = "Data was successfully exported from "& oNet.ComputerName &" for link '"& parm_linkName &"' at "& gl_oSupport.getMyTimeStamp()
'   Call gl_oSupport.sendMail (parm_emailFrom, parm_emailTo, msgText, "", msgText, "", parm_smtpServer, parm_smtpUser, parm_smtpPass, 25)
'   Set oNet = Nothing

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Sub


' #####
' ## runExportImport (ByVal in_rei_isRunNeeded)
Sub runExportImport (ByVal in_rei_isRunNeeded)
   Dim rei_ret

   Dim fn: fn = "runExportImport"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   ' !!! should not get here if AltTransport because of logic in isRunNeeded

   If parm_linkType = gl_icsLinkType.file Then
      Call log2file (" # "& fn &": Export & Import sync direction is not available for file mode ["& gl_linkExportFile &"]", "E")
   Else
      rei_ret = getFile (gl_linkExportFile, gl_linkImportFile, gl_linkDirName)

      If rei_ret <> True Then Call log2file (" # "& fn &": file is not accessible ["& gl_linkExportFile &"]", "W")

      Dim rei_statFileLoc: rei_statFileLoc = gl_oFso.BuildPath (gl_appSetupFilePath, gl_stateFile)

      If gl_oFso.FileExists (rei_statFileLoc) Then ' Not first time run
         If gl_oFso.FolderExists (gl_appSetupFilePath) And gl_oFso.GetFolder (gl_appSetupFilePath).Files.Count > 0 Then
            Call ics_update ("", gl_linkImportFile)

            Dim rei_cmpRet: rei_cmpRet = gl_oFileStruct.compareStructure (rei_statFileLoc)
            Dim rei_cmpOldRet: rei_cmpOldRet = gl_oFileStruct.compareStructure (gl_oFso.BuildPath (gl_appSetupFilePath &"\"& cOldStateDir, gl_stateFile))

            If (rei_cmpRet("status") < 0 Or rei_cmpOldRet("status") < 0) Then
               Call log2file (" # "& fn &": "& rei_cmpRet("description") &" ["& parm_linkName &"]", "E")
            
               If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
               Exit Sub
            End If

            ' !!! Call cleanExport, Failed to parse old structure
            Call runJob ("export", "RunAfter") ' !!! If everything is up today then it will not executed
         End If
      End If

      If parm_bBackUpRunExport Then Call runBackUp (False) ' was

      Call ics_unpack (gl_linkImportFile, gl_linkExportFile)
      Call runJob ("import", "RunAfter")

      If parm_bBackUpRunImport Then Call runBackUp (True) ' now

      Call createStateFile ("", gl_stateFileCurr)
      Call putFile (gl_linkImportFile, gl_linkExportFile, gl_linkDirName)
      Call putFile (gl_stateFileCurr, gl_stateFile, gl_linkDirName)
   End If

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Sub


' ### logic for importExport mode:

' * compare localCurr with remote
' * compare localCurr with local old state file localOld
' * rename localFiles if needed
' * update remote archive with local changes

' localOld <-- - new [newFilesOld] + -- localCurr -- + new [newFiles] - --> remote
' localOld <-- + del [delFilesOld] - -- localCurr -- - del [delFiles] + --> remote

' * modif remote & local
' rename oldest modified file name_modifTS.ext

' * for newFiles
' if localCurr -> localOld = exists then {
'    if (localOld date != localCurr date) then keep new file - file was removed on other PC, but updated on local PC
'    else DELETE FROM ARC - file removed on other PC
' }
' if localCurr -> localOld = newFilesOld then - file created on local PC

' * for delFiles
' if localCurr -> localOld = exists then {
'    if (localOld date < remote date) then - file crated on other PC
'    else DELETE FROM ARC - file removed on local PC
' }
' if localCurr -> localOld = not exsits then - file created on other PC

' * update remote archive


' #####
' ## ics_update (ByVal in_u_itemToUpdate, ByVal in_u_linkExportFile)
' ## return: true || false
Function ics_update (ByVal in_u_itemToUpdate, ByVal in_u_linkExportFile) ' return: true || false
   Dim u_parms, u_cmd, u_ret

   Dim fn: fn = "ics_update"
   Dim u_noRecursion: p_noRecursion = ""

   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   If parm_bNoTopLevelFiles Or parm_bNoTopLevelFolders Then u_noRecursion = "norecursion"
   If parm_bNoTopLevelFolders Then in_u_itemToPack = "*.*"

   ics_update = True
   
   u_parms = """"& gl_syncRootDirExpand  &""" """& in_u_linkExportFile &""" "&_
            """"& gl_appLoc &""" """& in_u_itemToPack &""" "& parm_syncPass &" """& gl_scriptWorkDir &""" "& parm_archType &" "& u_noRecursion
   u_cmd = """"& gl_appLoc &"\ics_update.bat"" "& u_parms
   u_ret = gl_oShell.Run (u_cmd &" >> """& gl_scriptOutLoc &"""", 0, true)

   ' copy output form ics_update to main out file
   u_parms = """"& gl_scriptOutLoc &""" """& gl_scriptWorkDir &"\ics_update.out"""
   u_cmd = """"& gl_appLoc &"\ics_merge.bat"" "& u_parms
   Call gl_oShell.Run (u_cmd, 0, true)

   If u_ret = 400 Then
      Call log2file ("Synchronize phase: Packing problem for "& in_u_linkExportFile, "E")
      ics_update = False
   ElseIf u_ret = 900 Then
      Call log2file ("Synchronize phase: Missing parameters", "W")
      ics_update = False
   End If

   ' copy ics_update.warr to log file
   Dim u_oWarrFile: Set u_oWarrFile = gl_oFso.GetFile (gl_scriptWorkDir &"\ics_update.warr")
   If u_oWarrFile.Size > 0 Then
      u_parms = """"& gl_appSetupFilePath &"\"& cAppLogFile &""" """& gl_scriptWorkDir &"\ics_update.warr"""
      u_cmd = """"& gl_appLoc &"\ics_merge.bat"" "& u_parms
      Call gl_oShell.Run (u_cmd, 0, true)
   End If
   Set u_oWarrFile = Nothing

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Function


' #####
' ## reportProblem ()
' ## return: 0 - ok || 1
Function reportProblem () ' return: 0 - ok || 1
   reportProblem = 0
   
   If Err Then
      Dim rp_msgType

      rp_msgType = "W"

      If Err.Number < 50000 Then
         rp_msgType = "E"
      End If

      Call log2file (Err.Number &": "& Err.Description, rp_msgType)

      Err.Clear
      reportProblem = 1
   End If
End Function


' #####
' ## setDefaultSetup ()
Sub setDefaultSetup ()
'   MsgBox "setDefaultSetup"

   gl_scriptErrBuff = ""
   If gl_tcpc_debug = True Then gl_tcpc_debugStack = ""

   parm_bLinkActive = False
   parm_linkName = ""
   parm_linkType = gl_icsLinkType.directory
   parm_dirDestRoot = "" '"%USERPROFILE%\Favorites"
   parm_ftp = ""
   parm_user = ""
   parm_pass = ""
   parm_smtpServer = ""
   parm_smtpUser = ""
   parm_smtpPass = ""
   parm_emailFrom = ""
   parm_emailTo = ""
   parm_syncType = gl_icsSyncType.exportOnly
   parm_syncPass = ""
   parm_bImportAltTransportUseFlag = False
   parm_importAltTransportPath = ""
   parm_bImportOnlyForceRefresh = False
   parm_bPackFileImportExport = False
   parm_exportMaxFileSize = 0
   parm_bBackUpRunImport = False
   parm_bBackUpRunExport = False
   parm_backUpDir = ""
   parm_backUpCount = 0
   parm_archType = ""
   parm_bExportImportFirstTime = False
   parm_bNoTopLevelFolders = False
   parm_bNoTopLevelFiles = False

   parm_aExportExcludePath = Array ()
   parm_aExportExcludeFile = Array ()
   parm_aExportIgnoreChangedFile = Array ()

   parm_aImportExcludePath = Array ()
   parm_aImportExcludeFile = Array ()

   parm_oFilesToProcess.RemoveAll
   parm_hImportJobToRunBefore.RemoveAll
   parm_hImportJobToRunAfter.RemoveAll
   parm_hExportJobToRunBefore.RemoveAll
   parm_hExportJobToRunAfter.RemoveAll

   gl_aFullyUsedDirs = Array ()
   gl_aFullyUsedFiles = Array ()

   gl_fLinkOwnerChanged = False
End Sub


' #####
' ## cleanUp ()
Sub cleanUp ()
   Dim cu_tempPath, cu_subFolder, cu_folderName, cu_inProcessFileLoc, cu_oInProcessFile
   Dim cu_timeDiff

   Dim fn: fn = "cleanUp"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   cu_tempPath = gl_oShell.ExpandEnvironmentStrings ("%TEMP%")

   Dim cu_oFolder: Set cu_oFolder = gl_oFso.GetFolder (cu_tempPath)
   Dim cu_colFolders: Set cu_colFolders = cu_oFolder.SubFolders

   Dim cu_oRegEx: Set cu_oRegEx = CreateObject("VBScript.RegExp")
   cu_oRegEx.IgnoreCase = True
   cu_oRegEx.Pattern = "^ckz_\w{4}"

   For Each cu_subFolder In cu_colFolders
      cu_folderName = Replace (LCase (cu_subFolder), LCase (cu_tempPath &"\"), "")

      If cu_oRegEx.Test (cu_folderName) Then
         cu_inProcessFileLoc = cu_subFolder &"\"& cfnInProcess

         If gl_oFso.FileExists (cu_inProcessFileLoc) Then
            Set cu_oInProcessFile = gl_oFso.GetFile (cu_inProcessFileLoc)
            cu_timeDiff = Abs (DateDiff ("n", Now, cu_oInProcessFile.DateLastModified)) ' h - Hour; n - Minute
            Set cu_oInProcessFile = Nothing

            If cu_timeDiff > cMaxRunTime Then Call gl_oFso.DeleteFolder (cu_subFolder, True)
         Else
            'Call gl_oFso.DeleteFolder (cu_subFolder, True)
            Call gl_oShell.Run ("cmd /c rmdir /S /Q "& cu_subFolder, 0, True)
         End If
      End If
   Next

   Set cu_oRegEx = Nothing
   Set cu_oFolder = Nothing
   Set cu_colFolders = Nothing
   Set cu_oInProcessFile = Nothing

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Sub


' #####
' ## getSetup ()
Sub getSetup ()
   Dim gs_oRegEx, gs_colMatches
   Dim gs_oStream
   Dim gs_oNodeTmpMain, gs_oNodeTmp, gs_oFileInfo
   Dim gs_tmp, gs_fn

   Dim fn: fn = "getSetup"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   On Error Resume Next

   If gl_oFso.FolderExists (gl_scriptWorkDir) Then 
      Call gl_oShell.Run ("cmd /c rmdir /S /Q "& gl_scriptWorkDir, 0, True)
   End If

   Call gl_oSupport.makeDir (gl_oFso, gl_scriptWorkDir)

   If Err Then
      On Error Goto 0
      Err.Description = "clean working dir"
      Err.Raise (19013)
   End If

   ' parm_linkName
   parm_linkName = oNode.selectSingleNode("./linkName").Text

   If parm_linkName = "" Then
      On Error Goto 0
      Err.Description = "We are sorry, but we can't complete process without valid link name"
      Err.Raise (19002)
   End If

   Set gs_oRegEx = CreateObject("VBScript.RegExp")
   gs_oRegEx.IgnoreCase = True
   gs_oRegEx.Pattern = "^\w+$"
   Set gs_colMatches = gs_oRegEx.Execute (parm_linkName)

   If gs_colMatches.Count < 1 Then
      On Error Goto 0
      Err.Description = "We are sorry, but we can't complete process without valid link name"
      Err.Raise (19003)
   End If

   ' parm_linkType
   If oNode.selectSingleNode("./linkType").Text <> "" Then parm_linkType = oNode.selectSingleNode("./linkType").Text

   ' parm_syncType
   If oNode.selectSingleNode("./syncType").Text <> "" Then parm_syncType = oNode.selectSingleNode("./syncType").Text

   ' parm_bLinkActive
   parm_bLinkActive = gl_oStringManip.getBoolValue (oNode.selectSingleNode("./linkActive").Text)
   If Not parm_bLinkActive Then
      gs_tmp = "Link is disabled  -- link: "& parm_linkName &", type: "& parm_linkType &", syncType: "& parm_syncType

      If parm_linkName <> gl_restoreLinkName Then
         On Error Goto 0
         Err.Description = gs_tmp
         Err.Raise (59014)
      End If

      Call log2file (gs_tmp, "W")
   End If

   ' parm_dirDestRoot
   parm_dirDestRoot = oNode.selectSingleNode("./syncRootDir").Text

   If parm_dirDestRoot = "" Then
      On Error Goto 0
      Err.Description = "We are sorry, but we can't complete process without valid SOURCE directory."
      Err.Raise (19007)
   End If

   gl_syncRootDirExpand = muExpandEnvironmentStrings (parm_dirDestRoot, null)

   ' parm_exportMaxFileSize
   parm_exportMaxFileSize = oNode.selectSingleNode ("./export/exportFile/maxFileSize").Text
   If parm_exportMaxFileSize = "" Then parm_exportMaxFileSize = 0

   ' parm_oFilesToProcess
   If parm_linkType = gl_icsLinkType.file Then
      Set gs_oNodeTmpMain = oNode.selectNodes ("./export/exportFile/files/file")

      For Each gs_oNodeTmp In gs_oNodeTmpMain
         gs_tmp = muExpandEnvironmentStrings (gs_oNodeTmp.Text, null)
         gs_tmp = Replace (gs_tmp, gl_syncRootDirExpand, "")

         If gs_tmp <> "" Then
            Set gs_oFileInfo = New CIcsFileInfo
            gs_oFileInfo.exportRequest = True
            gs_oFileInfo.importRequest = False
            gs_oFileInfo.fileName = ""

            parm_oFilesToProcess.Add gs_tmp, gs_oFileInfo
         End If
      Next

      Set gs_oNodeTmpMain = oNode.selectNodes ("./import/importFile/files/file")

      For Each gs_oNodeTmp In gs_oNodeTmpMain
         gs_tmp = muExpandEnvironmentStrings (gs_oNodeTmp.Text, null)
         gs_tmp = Replace (gs_tmp, gl_syncRootDirExpand, "")

         If gs_tmp <> "" Then
            If parm_oFilesToProcess.Exists(gs_tmp) Then
               parm_oFilesToProcess(gs_tmp).importRequest = True
            Else
               Set gs_oFileInfo = New CIcsFileInfo
               gs_oFileInfo.importRequest = True
               gs_oFileInfo.exportRequest = False
               gs_oFileInfo.fileName = ""

               parm_oFilesToProcess.Add gs_tmp, gs_oFileInfo
            End If
         End If
      Next

      ' parm_bPackFileImportExport
      parm_bPackFileImportExport = gl_oStringManip.getBoolValue (oNode.selectSingleNode ("./packFileImportExport").Text)
   End If

   ' parm_bImportAltTransportUseFlag
   parm_bImportAltTransportUseFlag = gl_oStringManip.getBoolValue (oNode.selectSingleNode("./import/importAltTransport/altTransportFlag").Text)

   ' parm_importAltTransportPath
   parm_importAltTransportPath = oNode.selectSingleNode("./import/importAltTransport/altTransportPath").Text

   If parm_bImportAltTransportUseFlag And parm_importAltTransportPath = "" Then
      On Error Goto 0
      Err.Description = "We are sorry, but we can't complete process without valid AltTransportPath."
      Err.Raise (19010)
   End If

   If Not parm_syncType = gl_icsSyncType.backupOnly Then
      ' parm_ftp
      parm_ftp = oNode.selectSingleNode("./linkFtp/host").Text
      
      If parm_ftp = "" Then
         On Error Goto 0
         Err.Description = "We are sorry, but we can't complete process without valid FTP host name."
         Err.Raise (19004)
      End If

      ' parm_user
      parm_user = oNode.selectSingleNode("./linkFtp/user").Text

      If parm_user = "" Then
         On Error Goto 0
         Err.Description = "We are sorry, but we can't complete process without valid USER name."
         Err.Raise (19005)
      End If

      ' parm_pass
      parm_pass = oNode.selectSingleNode("./linkFtp/pass").Text

      If parm_pass = "" Then
         On Error Goto 0
         Err.Description = "We are sorry, but we can't complete process without valid PASSWORD name."
         Err.Raise (19006)
      End If

      ' parm_smtpServer
      parm_smtpServer = oNode.selectSingleNode("./linkEmail/smtpServer").Text

      ' parm_smtpUser
      parm_smtpUser = oNode.selectSingleNode("./linkEmail/user").Text

      ' parm_smtpPass
      parm_smtpPass = oNode.selectSingleNode("./linkEmail/pass").Text

      ' parm_emailFrom
      parm_emailFrom = oNode.selectSingleNode("./linkEmail/emailFrom").Text

      ' parm_emailTo
      parm_emailTo = oNode.selectSingleNode("./export/emailToNotif").Text

      If (parm_emailTo <> "" And (parm_smtpServer = "" Or parm_smtpUser = "" Or parm_smtpPass = "" Or parm_emailFrom = "") ) Then
         On Error Goto 0
         Err.Description = "We are sorry, but we can't complete process without correct e-mail setup."
         Err.Raise (19008)
      End If
   End If

   ' parm_syncPass
   parm_syncPass = oNode.selectSingleNode("./syncPass").Text

   ' parm_backUpDir
   If oNode.selectSingleNode("./backUp/backUpDir").Text <> "" Then parm_backUpDir = oNode.selectSingleNode("./backUp/backUpDir").Text
   parm_backUpDir = muExpandEnvironmentStrings(parm_backUpDir, null)

   ' parm_backUpCount
   If oNode.selectSingleNode("./backUp/backUpCount").Text <> "" Then parm_backUpCount = oNode.selectSingleNode("./backUp/backUpCount").Text

   ' parm_archType
   If oNode.selectSingleNode("./packFormat").Text <> "" Then parm_archType = "-t"& oNode.selectSingleNode("./packFormat").Text

   ' parm_bImportOnlyForceRefresh
   If parm_syncType = gl_icsSyncType.importOnly Then
      parm_bImportOnlyForceRefresh = gl_oStringManip.getBoolValue (oNode.selectSingleNode("./import/importDirectory/forceRefresh").Text)
   End If

   ' parm_hImportJobToRunBefore
   Set gs_oNodeTmpMain = oNode.selectNodes("./import/jobsToRunBefore/job")
   For Each gs_oNodeTmp In gs_oNodeTmpMain
      If gs_oNodeTmp.Text <> "" Then
         gs_tmp = muExpandEnvironmentStrings (gs_oNodeTmp.selectSingleNode("./cmd").Text, null)
         gs_oRegEx.Pattern = "(^\s*"")(.*)(\s*""$)"
         gs_fn = gs_oRegEx.Replace (gs_tmp, "$2")

         If Not gl_oStringManip.getBoolValue (gs_oNodeTmp.selectSingleNode("./active").Text) Then
            Call log2file ("Job is currently disabled for importJobToRunBefore -- file: "& gs_tmp &", linkName: "& parm_linkName, "W")
         ElseIf gs_fn <> "" Then
            If Not gl_oFso.FileExists (gs_fn) Then Call log2file ("File does not exist for importJobToRunBefore -- file: "& gs_tmp &", linkName: "& parm_linkName, "W")

            If Not parm_hImportJobToRunBefore.Exists (gs_tmp) Then parm_hImportJobToRunBefore.Add gs_tmp, ""
         End If
      End If
   Next

   ' parm_hImportJobToRunAfter
   Set gs_oNodeTmpMain = oNode.selectNodes("./import/jobsToRunAfter/job")
   For Each gs_oNodeTmp In gs_oNodeTmpMain
      If gs_oNodeTmp.Text <> "" Then
         gs_tmp = muExpandEnvironmentStrings (gs_oNodeTmp.selectSingleNode("./cmd").Text, null)
         gs_oRegEx.Pattern = "(^\s*"")(.*)(\s*""$)"
         gs_fn = gs_oRegEx.Replace (gs_tmp, "$2")

         If Not gl_oStringManip.getBoolValue (gs_oNodeTmp.selectSingleNode("./active").Text) Then
            Call log2file ("Job is currently disabled for importJobToRunAfter -- file: "& gs_tmp &", linkName: "& parm_linkName, "W")
         ElseIf gs_fn <> "" Then
            If Not gl_oFso.FileExists (gs_fn) Then Call log2file ("File does not exist for importJobToRunAfter -- file: "& gs_tmp &", linkName: "& parm_linkName, "W")

            If Not parm_hImportJobToRunAfter.Exists (gs_tmp) Then parm_hImportJobToRunAfter.Add gs_tmp, ""
         End If
      End If
   Next

   ' parm_bBackUpRunImport
   parm_bBackUpRunImport = gl_oStringManip.getBoolValue (oNode.selectSingleNode("./import/backUpRun").Text)

   ' parm_bBackUpRunExport
   parm_bBackUpRunExport = gl_oStringManip.getBoolValue (oNode.selectSingleNode("./export/backUpRun").Text)

   If (parm_bBackUpRunImport Or parm_bBackUpRunExport) And parm_backUpDir = "" Then
      On Error Goto 0
      Err.Description = "We are sorry, but we can't complete process without valid BackUp setup."
      Err.Raise (19009)
   End if

   ' parm_hExportJobToRunBefore
   Set gs_oNodeTmpMain = oNode.selectNodes("./export/jobsToRunBefore/job")
   For Each gs_oNodeTmp In gs_oNodeTmpMain
      If gs_oNodeTmp.Text <> "" Then
         gs_tmp = muExpandEnvironmentStrings (gs_oNodeTmp.selectSingleNode("./cmd").Text, null)
         gs_oRegEx.Pattern = "(^\s*"")(.*)(\s*""$)"
         gs_fn = gs_oRegEx.Replace (gs_tmp, "$2")

         If Not gl_oStringManip.getBoolValue (gs_oNodeTmp.selectSingleNode("./active").Text) Then
            Call log2file ("Job is currently disabled for exportJobToRunBefore -- file: "& gs_tmp &", linkName: "& parm_linkName, "W")
         ElseIf gs_fn <> "" Then
            If Not gl_oFso.FileExists (gs_fn) Then Call log2file ("File does not exist for exportJobToRunBefore -- file: "& gs_tmp &", linkName: "& parm_linkName, "W")

            If Not parm_hExportJobToRunBefore.Exists (gs_tmp) Then parm_hExportJobToRunBefore.Add gs_tmp, ""
         End If
      End If
   Next

   ' parm_hExportJobToRunAfter
   Set gs_oNodeTmpMain = oNode.selectNodes("./export/jobsToRunAfter/job")
   For Each gs_oNodeTmp In gs_oNodeTmpMain
      If gs_oNodeTmp.Text <> "" Then
         gs_tmp = muExpandEnvironmentStrings (gs_oNodeTmp.selectSingleNode("./cmd").Text, null)
         gs_oRegEx.Pattern = "(^\s*"")(.*)(\s*""$)"
         gs_fn = gs_oRegEx.Replace (gs_tmp, "$2")

         If Not gl_oStringManip.getBoolValue (gs_oNodeTmp.selectSingleNode("./active").Text) Then
            Call log2file ("Job is currently disabled for exportJobToRunAfter -- file: "& gs_tmp &", linkName: "& parm_linkName, "W")
         ElseIf gs_fn <> "" Then
            If Not gl_oFso.FileExists (gs_fn) Then Call log2file ("File does not exist for exportJobToRunAfter -- file: "& gs_tmp &", linkName: "& parm_linkName, "W")

            If Not parm_hExportJobToRunAfter.Exists (gs_tmp) Then parm_hExportJobToRunAfter.Add gs_tmp, ""
         End If
      End If
   Next

   ' ExportExclude
   Set gs_oStream = gl_oFso.CreateTextFile (gl_packExcludeLoc &"Tmp", True, True)

   Set gs_oNodeTmpMain = oNode.selectNodes ("./export/exportDirectory/excludeDir/dir")
   For Each gs_oNodeTmp In gs_oNodeTmpMain
      gs_tmp = muExpandEnvironmentStrings (gs_oNodeTmp.Text, null)
      gs_tmp = Replace (gs_tmp, gl_syncRootDirExpand, "")

      If gs_tmp <> "" Then
         If gs_tmp = cNoTopLevelFolders Then ' a arc.zip *.* [no -r switch]
            parm_bNoTopLevelFolders = True
         Else
            gs_oStream.WriteLine gs_tmp
         End If

         Call gl_oSupport.arrayPush (parm_aExportExcludePath, gs_tmp)
      End If
   Next

   Set gs_oNodeTmpMain = oNode.selectNodes ("./export/exportDirectory/excludeFile/file")
   For Each gs_oNodeTmp In gs_oNodeTmpMain
      gs_tmp = muExpandEnvironmentStrings (gs_oNodeTmp.Text, null)
      gs_tmp = Replace (gs_tmp, gl_syncRootDirExpand, "")

      If gs_tmp <> "" Then
         If gs_tmp = cNoTopLevelFiles Then ' a arc.zip -x!*.* [no -r switch]
            parm_bNoTopLevelFiles = True
            gs_oStream.WriteLine "*.*" ' !!! for 7z to exclude top level files
         Else
            gs_oStream.WriteLine gs_tmp
         End If

         Call gl_oSupport.arrayPush (parm_aExportExcludeFile, gs_tmp)
      End If
   Next
   gs_oStream.close

   Call muChangeFileEncoding (gl_packExcludeLoc &"Tmp", gl_packExcludeLoc, "utf-8") ' 7z can't read unicode list files & has utf-8 as default charset

   ' ImportExclude
   Set gs_oStream = gl_oFso.CreateTextFile (gl_unpackIgnoreLoc &"Tmp", True, True) ' ***

   Set gs_oNodeTmpMain = oNode.selectNodes ("./import/importDirectory/excludeDir/dir")
   For Each gs_oNodeTmp In gs_oNodeTmpMain
      gs_tmp = muExpandEnvironmentStrings (gs_oNodeTmp.Text, null)
      gs_tmp = Replace (gs_tmp, gl_syncRootDirExpand, "")

      If gs_tmp <> "" Then
         Call gl_oSupport.arrayPush (parm_aImportExcludePath, gs_tmp)
         gs_oStream.WriteLine gs_tmp ' ***
      End If
   Next

   Set gs_oNodeTmpMain = oNode.selectNodes ("./import/importDirectory/excludeFile/file")
   For Each gs_oNodeTmp In gs_oNodeTmpMain
      gs_tmp = muExpandEnvironmentStrings (gs_oNodeTmp.Text, null)
      gs_tmp = Replace (gs_tmp, gl_syncRootDirExpand, "")

      If gs_tmp <> "" Then
         Call gl_oSupport.arrayPush (parm_aImportExcludeFile, gs_tmp)
         gs_oStream.WriteLine gs_tmp ' ***
      End If
   Next

   gs_oStream.close ' ***

   Call muChangeFileEncoding (gl_unpackIgnoreLoc &"Tmp", gl_unpackIgnoreLoc, "utf-8") ' *** 7z can't read unicode list files & has utf-8 as default charset

   ' parm_aExportIgnoreChangedFile
   Set gs_oNodeTmpMain = oNode.selectNodes ("./export/exportDirectory/ignoreFile/file")
   For Each gs_oNodeTmp In gs_oNodeTmpMain
      gs_tmp = muExpandEnvironmentStrings (gs_oNodeTmp.Text, null)
      gs_tmp = Replace (gs_tmp, gl_syncRootDirExpand, "")

      If gs_tmp <> "" Then
         Call gl_oSupport.arrayPush (parm_aExportIgnoreChangedFile, gs_tmp)
      End If
   Next

   Set gs_oStream = Nothing
   Set gs_oNodeTmpMain = Nothing
   Set gs_oNodeTmp = Nothing
   Set gs_oRegEx = Nothing

   If parm_syncPass = "" Then parm_syncPass = """"""

   Err.Clear
   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Sub


' #####
' ## createStateFile (ByVal in_csf_item, ByVal in_csf_stateFileName)
' ## return: file count
Function createStateFile (ByVal in_csf_item, ByVal in_csf_stateFileName) ' return: file count
   Dim fn: fn = "createStateFile:Fatal error to access """& gl_syncRootDirExpand &"\"& in_csf_item &""""
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   Call gl_oFileStruct.setStartingPoint (gl_syncRootDirExpand)

   If parm_syncType = gl_icsSyncType.importOnly Then
      createStateFile = gl_oFileStruct.populateStructure (in_csf_item, parm_aImportExcludePath, parm_aImportExcludeFile, "")
   Else
      createStateFile = gl_oFileStruct.populateStructure (in_csf_item, parm_aExportExcludePath, parm_aExportExcludeFile, "")
   End If

   Call gl_oFileStruct.saveStructure (gl_scriptWorkDir &"\"& in_csf_stateFileName)

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Function


' #####
' ## getFile (ByVal in_gf_remoteName, ByVal in_gf_localName, ByVal in_gf_linkDirName)
' ## return: true || false
Function getFile (ByVal in_gf_remoteName, ByVal in_gf_localName, ByVal in_gf_linkDirName) ' return: true || false
   Dim gf_linkRemotePath, gf_ret

   Dim fn: fn = "getFile"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   getFile = False
   
   gf_linkRemotePath = gl_linkRemoteRoot
   If in_gf_linkDirName <> "" Then  gf_linkRemotePath = gf_linkRemotePath &"/"& in_gf_linkDirName

   If gl_restoreLinkName <> "" Then
      gl_oFso.CopyFile parm_backUpDir &"\"& gl_restoreBackUpFileName, gl_scriptWorkDir &"\"& in_gf_localName
      getFile = True
   ElseIf parm_bImportAltTransportUseFlag Then
      getFile = muSaveFileFromWeb (gf_linkRemotePath &"/"& in_gf_remoteName, gl_scriptWorkDir &"\"& in_gf_localName)
   Else
      gf_ret = ics_get (in_gf_remoteName, gf_linkRemotePath, in_gf_localName)

      If gf_ret = 200 Then
         Err.Description = " # "& fn &": Cannot connect to FTP"
         Err.Raise (gf_ret)
      ElseIf gf_ret = 210 Then
         Err.Description = " # "& fn &": Check FTP Settings"
         Err.Raise (gf_ret)
      ElseIf gf_ret = 900 Then
         Err.Description = " # "& fn &": Missing parameters"
         Err.Raise (gf_ret)
      ElseIf gf_ret = 220 Then
         Call log2file (" # "& fn &": No such file or directory -- "& gf_linkRemotePath &"/"& in_gf_remoteName &" @ "& parm_ftp, "W")
      Else
         getFile = True
      End If
   End If

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Function


' #####
' ## ics_get (ByVal in_g_remoteName, ByVal in_g_remotePath, ByVal in_g_localName)
' ## return: errorCode
Function ics_get (ByVal in_g_remoteName, ByVal in_g_remotePath, ByVal in_g_localName) ' return: errorCode
   Dim g_parms, g_cmd

   Dim fn: fn = "ics_get"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   ics_get = 999

   g_parms = parm_ftp &" "& parm_user &" "& parm_pass &" """& in_g_remoteName &""" """& in_g_remotePath &""" """& in_g_localName &""" """& gl_scriptWorkDir &""""
   g_cmd = """"& gl_appLoc &"\ics_get.bat"" "& g_parms
   ics_get = gl_oShell.Run (g_cmd &" >> """& gl_scriptOutLoc &""" """, 0, true)

   ' copy output form ics_get to main out file
   g_parms = """"& gl_scriptOutLoc &""" """& gl_scriptWorkDir &"\ics_get.out"""
   g_cmd = """"& gl_appLoc &"\ics_merge.bat"" "& g_parms
   Call gl_oShell.Run (g_cmd, 0, true)   

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Function


' #####
' ## isRunNeeded ()
' ## return: 0 then no need; 1 - run export; 2 - run all;
Function isRunNeeded () ' if 0 then no need; 1 - run export; 2 - run all;
   Dim fn: fn = "isRunNeeded"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   isRunNeeded = 2

   If gl_restoreLinkName <> "" Then
      If parm_linkName <> gl_restoreLinkName Then
         Call log2file (" # "& fn &": Restore mode -- skip "& parm_linkName &" link", "I")

         isRunNeeded = 0
      Else
         parm_syncType = gl_icsSyncType.importOnly
         parm_bImportOnlyForceRefresh = True
         parm_bBackUpRunImport = False
         
         Call log2file (" # "& fn &": Restore mode -- run "& parm_linkName &" link", "I")

         If parm_linkType = gl_icsLinkType.file Then
            Dim irn_oFileInfo: Set irn_oFileInfo = New CIcsFileInfo
            irn_oFileInfo.exportRequest = False
            irn_oFileInfo.importRequest = True
            irn_oFileInfo.fileName = gl_restoreFileName

            Set parm_oFilesToProcess.Item (gl_restoreFileName) = irn_oFileInfo
         Else
            Dim irn_aRestoreExcludePath: irn_aRestoreExcludePath = Array ()
            Dim irn_aRestoreExcludeFile: irn_aRestoreExcludeFile = Array ()

            Call gl_oSupport.arrayPush (irn_aRestoreExcludePath, parm_aExportExcludePath)
            Call gl_oSupport.arrayPush (irn_aRestoreExcludeFile, parm_aExportExcludeFile)

            Call gl_oSupport.arrayPush (irn_aRestoreExcludePath, parm_aImportExcludePath)
            Call gl_oSupport.arrayPush (irn_aRestoreExcludeFile, parm_aImportExcludeFile)

            Call gl_oFileStruct.setStartingPoint (gl_syncRootDirExpand)
            Call gl_oFileStruct.populateStructure ("", irn_aRestoreExcludePath, irn_aRestoreExcludeFile, "")
            gl_aFullyUsedDirs = gl_oFileStruct.getFullyUsedFolders ()
            gl_aFullyUsedFiles = gl_oFileStruct.getFullyUsedFiles ()
         End If
      End If
       
      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
      Exit Function
   End If

   If parm_syncType <> gl_icsSyncType.importOnly Then
      Call runJob ("export", "RunBefore")
   ElseIf parm_syncType = gl_icsSyncType.importOnly Or parm_syncType = gl_icsSyncType.synchronize Then
      Call runJob ("import", "RunBefore")
   End If

   If checkForOwnerChanges () Then ' No need to run 
      isRunNeeded = 0
   ElseIf parm_linkType = gl_icsLinkType.file Then
      isRunNeeded = isRunNeeded4File ()
   ElseIf gl_linkDirName <> "" Then ' gl_icsLinkType.directory
      isRunNeeded = isRunNeeded4Dir ()
   Else
      isRunNeeded = 0      
      Call log2file ("Use of AltTransport in "& gl_icsLinkType.directory &" mode is not allowed.", "E")
   End If

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)

End Function


' #####
' ## isRunNeeded4File ()
' ## return: 0 then no need; 1 - run export; 2 - run all;
Function isRunNeeded4File () ' return: 0 then no need; 1 - run export; 2 - run all;
   Dim irn4f_item, irn4f_retIcsGet, irn4f_stateOld

   Dim fn: fn = "isRunNeeded4File"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   isRunNeeded4File = 2

   If parm_oFilesToProcess.Count = 0 Then
      Call log2file (" # "& fn &": no files specified to work with", "I")
      isRunNeeded4File = 0

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
      Exit Function
   End If

   For Each irn4f_item In parm_oFilesToProcess.Keys
      irn4f_retIcsGet = False

      If parm_syncType = gl_icsSyncType.backupOnly Then
         irn4f_retIcsGet = True
         irn4f_stateOld = parm_backUpDir &"\"& irn4f_item & cStateExt
      Else
         If Not parm_bPackFileImportExport And InStr (irn4f_item, gl_toPackExt) + 2 = Len (irn4f_item) Then
            irn4f_stateOld = gl_oStringManip.replaceRegExp (irn4f_item, gl_toPackExt, "") & cStateExt
         Else
            irn4f_stateOld = irn4f_item & cStateExt
         End If

         If gl_linkDirName = "" Then
            irn4f_retIcsGet = getFile (irn4f_item, irn4f_item & cfnImportedSuff, "") ' !!! important for runImport
         Else
            irn4f_retIcsGet = getFile (irn4f_stateOld, irn4f_stateOld, gl_linkDirName)
         End if

         irn4f_stateOld = gl_scriptWorkDir &"\"& irn4f_stateOld
      End If

      If irn4f_retIcsGet = True And Not parm_bPackFileImportExport And InStr (irn4f_item, gl_toPackExt) + 2 = Len (irn4f_item) Then
         Call gl_oFso.MoveFile (irn4f_stateOld, gl_scriptWorkDir &"\"& irn4f_item & cStateExt)
      End If

      If irn4f_retIcsGet And gl_linkDirName = "" Then ' this is separate "if" because logic to get & compare is little different for importAltTransport
         If ics_checkState (gl_syncRootDirExpand &"\"& irn4f_item, irn4f_item & cfnImportedSuff) = 0 Then
            Call log2file (" # "& fn &": file up today ["& irn4f_item &"], there is nothing to export or import file", "I")
            parm_oFilesToProcess.Remove (irn4f_item)
         Else
            Call log2file (" # "& fn &": sync -- "& irn4f_item, "I")
         End If
      ElseIf gl_oFso.FileExists (gl_syncRootDirExpand &"\"& irn4f_item) Then
         Call createStateFile (irn4f_item, irn4f_item & cCurrStateExt)

         If irn4f_retIcsGet = True Then
            Dim irn4f_fileCnt: irn4f_fileCnt = cmpStateFiles (irn4f_stateOld)

            If irn4f_fileCnt < 0 Then
               Call ics_rename (irn4f_item & cStateExt)

               Call log2file (" # "& fn &": Failed to parse state file, there is nothing to import", "E")
               parm_oFilesToProcess(irn4f_item).importRequest = False
            ElseIf irn4f_fileCnt = 0 Then
               Call log2file (" # "& fn &": file up today ["& irn4f_item &"], there is nothing to export or import", "I")
               parm_oFilesToProcess.Remove (irn4f_item)
            Else
               Call log2file (" # "& fn &": sync -- "& irn4f_item, "I")
            End If
         Else
            If parm_syncType = gl_icsSyncType.importOnly Then
               Call log2file (" # "& fn &": there is nothing to import ["& irn4f_item &"]", "I")
               parm_oFilesToProcess.Remove (irn4f_item)
            Else
               Call log2file (" # "& fn &": there is no need to run import ["& irn4f_item &"]", "I")

               parm_oFilesToProcess(irn4f_item).importRequest = False
            End If
         End If
      Else
         Call log2file (" # "& fn &": file does not exist ["& irn4f_item &"]", "W")

         parm_oFilesToProcess(irn4f_item).exportRequest = False

         If irn4f_retIcsGet = True Then
            If parm_syncType = gl_icsSyncType.importOnly Or parm_syncType = gl_icsSyncType.synchronize Then

               If parm_oFilesToProcess(irn4f_item).importRequest = False Then
                  Call log2file (" # "& fn &": there is nothing to export or import, ignore file ["& irn4f_item &"]", "I")
                  parm_oFilesToProcess.Remove (irn4f_item)
               Else
                  Call log2file (" # "& fn &": there is nothing to export ["& irn4f_item &"]", "I")
                  parm_oFilesToProcess(irn4f_item).exportRequest = False
               End If
            Else
               Call log2file (" # "& fn &": there is nothing to export, ignore file ["& irn4f_item &"]", "I")
               parm_oFilesToProcess.Remove (irn4f_item)
            End If
         Else
            Call log2file (" # "& fn &": there is nothing to export or import, problem to retrieve file ["& irn4f_item &"]", "W")
            parm_oFilesToProcess.Remove (irn4f_item)
         End If
      End If
   Next

   If parm_oFilesToProcess.Count = 0 Then isRunNeeded4File = 0
   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)

End Function


' #####
' ## isRunNeeded4Dir ()
' ## return: 0 then no need; 1 - run export; 2 - run all;
Function isRunNeeded4Dir () ' return: 0 then no need; 1 - run export; 2 - run all;
   Dim irn4d_stateOld, irn4d_retIcsGet, irn4d_stateCurr_fileCnt

   Dim fn: fn = "isRunNeeded4Dir"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   isRunNeeded4Dir = 2

   irn4d_stateOld = gl_stateFile

   ' !!! - What if inside no files only empty folders
   If (Not gl_oFso.FolderExists (gl_syncRootDirExpand)) _
      And (parm_syncType = gl_icsSyncType.backupOnly Or parm_syncType = gl_icsSyncType.exportOnly) Then

      Call log2file (" # "& fn &": there is nothing to export for link because folder does Not Exist", "I")
      isRunNeeded4Dir = 0

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
      Exit Function
   End If

   If parm_syncType = gl_icsSyncType.backupOnly Then
      irn4d_retIcsGet = True
      irn4d_stateOld = parm_backUpDir &"\"& gl_stateFile
   Else
      irn4d_retIcsGet = getFile (irn4d_stateOld, irn4d_stateOld, gl_linkDirName)
      irn4d_stateOld = gl_scriptWorkDir &"\"& gl_stateFile
   End If

   irn4d_stateCurr_fileCnt = createStateFile ("", gl_stateFileCurr)
   gl_aFullyUsedDirs = gl_oFileStruct.getFullyUsedFolders ()
   gl_aFullyUsedFiles = gl_oFileStruct.getFullyUsedFiles ()

   If irn4d_retIcsGet = False Then
      If parm_syncType = gl_icsSyncType.importOnly Then
         Call log2file (" # "& fn &": there is nothing to import for link", "I")
         isRunNeeded4Dir = 0

         If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
         Exit Function
      ElseIf parm_syncType = gl_icsSyncType.exportOnly Or parm_syncType = gl_icsSyncType.synchronize Then
         If irn4d_stateCurr_fileCnt > 0 Then
            Call log2file (" # "& fn &": there is nothing to import for link, "& parm_syncType, "I")
            isRunNeeded4Dir = 1

            If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
            Exit Function
         Else
            Call log2file (" # "& fn &": there is nothing to import and nothing to export for link, "& parm_syncType, "W")
            isRunNeeded4Dir = 0

            If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
            Exit Function
         End If
      End If
   End If

   If irn4d_stateCurr_fileCnt = 0 Then
      If parm_syncType <> gl_icsSyncType.importOnly Then
         Dim irn4d_syncMsg: irn4d_syncMsg = ""

         If parm_syncType = gl_icsSyncType.synchronize Then
            parm_syncType = gl_icsSyncType.importOnly
            isRunNeeded4Dir = 2

            irn4d_syncMsg = ", Sync type changed to "& parm_syncType
            ' !!! move state file to local state folder
         Else 
            isRunNeeded4Dir = 0
         End If

         Call log2file (" # "& fn &": there is nothing to export for link, "& parm_syncType & irn4d_syncMsg, "I")

         If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
         Exit Function
      End If

      If gl_oFso.FolderExists (gl_syncRootDirExpand) Then
         Dim irn4d_oFolder: Set irn4d_oFolder = gl_oFso.GetFolder (gl_syncRootDirExpand)

         If irn4d_oFolder.Files.Count > 0 Then ' !!! - What if inside no files only empty folders
            Call log2file (" # "& fn &": no update will be done for link", "E")
            isRunNeeded4Dir = 0

            If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
            Exit Function
         End If

         Set irn4d_oFolder = Nothing
      End If
   End If

   Dim irn4d_fileCnt: irn4d_fileCnt = cmpStateFiles (irn4d_stateOld)
   If irn4d_fileCnt < 0 Then
      Call ics_rename (gl_stateFile)

      Call log2file (" # "& fn &": Failed to parse state file, there is nothing to import", "E")
      isRunNeeded4Dir = 1
   ElseIf irn4d_fileCnt = 0 Then
      isRunNeeded4Dir = 0
   End If
   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)

End Function


' #####
' ## checkForOwnerChanges ()
' ## return: true || false
Function checkForOwnerChanges () ' return: true || false
   Dim cfoc_remoteName, cfoc_oXml, cfoc_ret

   checkForOwnerChanges = False


   If parm_syncType = gl_icsSyncType.backupOnly Then
      Exit Function
   ElseIf gl_linkDirName = "" Then
      Call log2file (" # No need to parse config - AltTransport", "I")

      Exit Function
   End If

   
   Dim fn: fn = "checkForOwnerChanges"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   
   cfoc_remoteName = gl_linkDirName & cConfigFileExt
   cfoc_ret = getFile (cfoc_remoteName, cfoc_remoteName, "")

   If cfoc_ret = False And parm_syncType = gl_icsSyncType.importOnly Then ' to catch ics_get 220 error code, for any ftp modes !!!
      Call log2file ("Networking problem or there is nothing to import because owner didn't export anything yet", "W")
      checkForOwnerChanges = True
      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)

      Exit Function
   End If


   Set cfoc_oXml = WScript.CreateObject ("Microsoft.XMLDOM")
   cfoc_oXml.async = "false"
   cfoc_oXml.Load (gl_scriptWorkDir &"\"& cfoc_remoteName)

   If cfoc_oXml.parseError.errorCode Then
      If parm_syncType = gl_icsSyncType.importOnly Then
         Err.Description = " # "& fn &": Failed to parse config"
         Err.Raise (19018)
      End If

      Call log2file (" # "& fn &": Failed to parse config", "W")
   ElseIf parm_syncType = gl_icsSyncType.importOnly Then
      Dim cfoc_syncPassChangeCount: cfoc_syncPassChangeCount = ""
      Dim cfoc_syncPassChangeCountOwner: cfoc_syncPassChangeCountOwner = ""
      Dim cfoc_archTypeOwner: cfoc_archTypeOwner = ""
      Dim cfoc_bPackFileImportExportOwner: cfoc_bPackFileImportExportOwner = False

      If oNode.selectSingleNode("./iter") Is Nothing Then
         cfoc_syncPassChangeCount = 0
      Else 
         cfoc_syncPassChangeCount = Int (oNode.selectSingleNode("./iter").Text)
      End If

      If cfoc_oXml.selectSingleNode("//iter") Is Nothing Then
         cfoc_syncPassChangeCountOwner = 0
      Else 
         cfoc_syncPassChangeCountOwner = Int (cfoc_oXml.selectSingleNode("//iter").Text)
      End If
      
      cfoc_archTypeOwner = cfoc_oXml.selectSingleNode("//packFormat").Text
      cfoc_bPackFileImportExportOwner = gl_oStringManip.getBoolValue (cfoc_oXml.selectSingleNode("//packFileImportExport").Text)

      If (cfoc_syncPassChangeCountOwner <> cfoc_syncPassChangeCount) Then
         Call log2file ("New password is required because link owner changed it.", "E")
         checkForOwnerChanges = True
      ElseIf (cfoc_archTypeOwner <> parm_archType) Then
         Call log2file ("New Pack/Unpack format is required because link owner changed it.", "E")
         checkForOwnerChanges = True
      ElseIf (parm_linkType = gl_icsLinkType.file And cfoc_bPackFileImportExportOwner <> parm_bPackFileImportExport) Then
         Call log2file ("Link owner changed the way files are exporting.", "E")            
         checkForOwnerChanges = True
      End If
   ElseIf parm_syncType = gl_icsSyncType.exportOnly Then
      Dim cfoc_ownerHost: cfoc_ownerHost = cfoc_oXml.selectSingleNode ("//hostName").Text
      Dim cfoc_ownerUser: cfoc_ownerUser = cfoc_oXml.selectSingleNode ("//hostUserId").Text

      Dim cfoc_colUserCurr, cfoc_aUserCurr
      cfoc_ret = gl_oSupport.getCurrUserId (cfoc_colUserCurr)
      cfoc_aUserCurr = cfoc_colUserCurr.Keys

      If gl_oSupport.getCompName () <> cfoc_ownerHost Or UCase (cfoc_aUserCurr(0)) <> cfoc_ownerUser Then
         Dim cfoc_oXmlLinkOwnerChange: Set cfoc_oXmlLinkOwnerChange = WScript.CreateObject("Microsoft.XMLDOM")
         Dim cfoc_oOwnerHostPrev: Set cfoc_oOwnerHostPrev = Nothing
         Dim cfoc_oOwnerUserPrev: Set cfoc_oOwnerUserPrev = Nothing
         Dim cfoc_bOwnerChanged: cfoc_bOwnerChanged = False
         Dim cfoc_oNodeLinkOwnerChange

         
         cfoc_oXmlLinkOwnerChange.async = "false"
         cfoc_oXmlLinkOwnerChange.Load (gl_oFso.BuildPath (gl_appSetupFilePath, cLinkOwnerChangeFile))

         If Not cfoc_oXmlLinkOwnerChange.parseError.errorCode Then
            Set cfoc_oNodeLinkOwnerChange = cfoc_oXmlLinkOwnerChange.selectSingleNode("//linkOwnerChanges/"& UCase (gl_linkDirName))

            Set cfoc_oOwnerHostPrev = cfoc_oNodeLinkOwnerChange.selectSingleNode ("./linkOwnerHostPrev")
            Set cfoc_oOwnerUserPrev = cfoc_oNodeLinkOwnerChange.selectSingleNode ("./linkOwnerUserPrev")
         End If


         If cfoc_oOwnerHostPrev Is Nothing Or cfoc_oOwnerUserPrev Is Nothing Then
            cfoc_bOwnerChanged = True
         ElseIf cfoc_oOwnerHostPrev.Text <> cfoc_ownerHost Or cfoc_oOwnerUserPrev.Text <> cfoc_ownerUser Then
            cfoc_bOwnerChanged = True
         End If

         If cfoc_bOwnerChanged = False Then
            Call log2file ("Link's ownership will be taking from host: "& cfoc_ownerHost &" and user: "& cfoc_ownerUser &" if sync required.", "W")

            gl_fLinkOwnerChanged = True
         Else
            Call log2file ("Other user took link's ownership and nothing will be exported from this PC. Currently ownership belongs to host: "& cfoc_ownerHost &" and user: "& cfoc_ownerUser, "E")

            checkForOwnerChanges = True
         End If

         Set cfoc_oOwnerHostPrev = Nothing
         Set cfoc_oOwnerUserPrev = Nothing
         Set cfoc_oNodeLinkOwnerChange = Nothing
         Set cfoc_oXmlLinkOwnerChange = Nothing
      End If
   End If

   Set cfoc_oXml = Nothing

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Function


' #####
' ## ics_checkState (ByVal in_cs_nameOld, ByVal in_cs_nameCurr)
' ## return: non-zero if files different
Function ics_checkState (ByVal in_cs_nameOld, ByVal in_cs_nameCurr) ' return: non-zero if files different -- Compares files as ASCII text
   Dim cs_parms, cs_cmd
   Dim cs_ret

   Dim fn: fn = "ics_checkState"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   cs_parms = """"& in_cs_nameOld &""" """& in_cs_nameCurr &""" """& gl_scriptWorkDir &""""
   cs_cmd = """"& gl_appLoc &"\ics_checkState.bat"" "& cs_parms
   cs_ret = gl_oShell.Run (cs_cmd &" >> """& gl_scriptOutLoc &"""", 0, true)

   If cs_ret = 900 Then
      Err.Description = "State phase: Missing parameters"
      Err.Raise (cs_ret)
   End If

   ics_checkState = cs_ret
   
   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Function


' #####
' ## cmpStateFiles (ByVal in_csf_nameOld)
' ## return: count of changes || -2 - failed to parse
Function cmpStateFiles (ByVal in_csf_nameOld) ' return: count of changes || -2 - failed to parse
   Dim fn: fn = "cmpStateFiles"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   Dim csf_cmpRet: Set csf_cmpRet = gl_oFileStruct.compareStructure (in_csf_nameOld)

   If (csf_cmpRet("status") = -2) Then
      Call log2file (" # "& fn &": "& csf_cmpRet("description"), "W")
      csf_cmpRet("status") = -2
   ElseIf (csf_cmpRet("status") = -1) Then
      If parm_syncType = gl_icsSyncType.backupOnly Or parm_syncType = gl_icsSyncType.exportOnly Then
         Call log2file (" # "& fn &": "& csf_cmpRet("description"), "W")
         csf_cmpRet("status") = 1
      End If
   ElseIf (csf_cmpRet("status") < -1) Then
      Call log2file (" # "& fn &": "& csf_cmpRet("description"), "E")
      csf_cmpRet("status") = 0
   ElseIf (csf_cmpRet("status") > 0) Then
      Dim csf_oReIgnoreSearch: Set csf_oReIgnoreSearch = CreateObject ("VBScript.RegExp"): csf_oReIgnoreSearch.IgnoreCase = True
      Dim csf_hModifFileToIgnore: Set csf_hModifFileToIgnore = CreateObject("Scripting.Dictionary")
      Dim csf_file, csf_fileIgnore

      For Each csf_file In csf_cmpRet("modifFiles")      
         For Each csf_fileIgnore In parm_aExportIgnoreChangedFile
            csf_fileIgnore = muExpandEnvironmentStrings (csf_fileIgnore, gl_scriptErrBuff)
            csf_fileIgnore = gl_oStringManip.escapeString (csf_fileIgnore)
            csf_fileIgnore = Replace (csf_fileIgnore, "*", ".*")

            csf_oReIgnoreSearch.Pattern = csf_fileIgnore &"$"

            If csf_oReIgnoreSearch.Execute (csf_file).Count > 0 And Not csf_hModifFileToIgnore.Exists (csf_file) Then
               csf_hModifFileToIgnore.Add csf_file, ""
            End If
         Next
      Next

      Dim modifMaxIndx: modifMaxIndx = UBound (csf_cmpRet("modifFiles"))
      Dim newMaxIndx: newMaxIndx = UBound (csf_cmpRet("newFiles"))
      Dim delMaxIndx: delMaxIndx = UBound (csf_cmpRet("delFiles"))
      Dim exclMaxIndx: exclMaxIndx = UBound (csf_cmpRet("exclFiles"))

      Call log2file (" # "& fn &": counts -- "& modifMaxIndx + 1 &" "& newMaxIndx + 1 &" "& delMaxIndx + 1 &" "& exclMaxIndx + 1, "I")
      Call log2file ("  Modified files: "& Join (csf_cmpRet("modifFiles"), vbcrlf), "I")
      If parm_syncType <> gl_icsSyncType.importOnly Then Call log2file ("  Export only ignored files: "& Join (parm_aExportIgnoreChangedFile, vbcrlf), "I")

      If csf_hModifFileToIgnore.Count = modifMaxIndx + 1 And newMaxIndx = -1 And delMaxIndx = -1 Then
         csf_cmpRet("status") = 0
      End If

      Call log2file ("  New files: "& Join (csf_cmpRet("newFiles"), vbcrlf), "I")
      Call log2file ("  Del files: "& Join (csf_cmpRet("delFiles"), vbcrlf), "I")
   End If

   cmpStateFiles = csf_cmpRet("status")
   
   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Function


' #####
' ## ics_unpack (ByVal in_up_file, ByVal in_up_fileOrigName)
Sub ics_unpack (ByVal in_up_file, ByVal in_up_fileOrigName)
   Dim up_parms, up_cmd, up_fileName, up_ret
   Dim up_forceRefresh: up_forceRefresh = "false"

   Dim fn: fn = "ics_unpack"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   If parm_bImportOnlyForceRefresh Then Call renameFullyUsedFiles (true)

   up_parms = """"& in_up_file &""" """& gl_syncRootDirExpand &""" "&_
            """"& gl_appLoc &""" "& up_forceRefresh &" """" "& parm_syncPass &" """& gl_scriptWorkDir &""""
   up_cmd = """"& gl_appLoc &"\ics_unpack.bat"" "& up_parms
   up_ret = gl_oShell.Run (up_cmd &" >> """& gl_scriptOutLoc &"""", 0, true)

   If up_ret = 300 Then
      If parm_bImportOnlyForceRefresh Then
         Call cleanDestRoot (false)
         Call renameFullyUsedFiles (false)
      End If

      Err.Description = "Unpack phase: arch"
      Err.Raise (up_ret)
   ElseIf up_ret = 310 Then
      If parm_bImportOnlyForceRefresh Then
         Call cleanDestRoot (false)
         Call renameFullyUsedFiles (false)
      End If

      Err.Description = "Unpack phase: import"
      Err.Raise (up_ret)
   ElseIf up_ret = 320 Then
      If parm_bImportOnlyForceRefresh Then
         Call cleanDestRoot (false)
         Call renameFullyUsedFiles (false)
      End If

      up_fileName = Replace (in_up_fileOrigName, gl_toPackExt, "")

      Call ics_rename (in_up_fileOrigName)
      Call ics_rename (up_fileName & cStateExt)

      Err.Description =  "Unpack phase: back up file corrupted! File marked as damaged and nothing has been imported." ' incorrect password: exit code for 7z.exe is 2 !!! 
      Err.Raise (19320)
   ElseIf up_ret = 330 Then
      If parm_bImportOnlyForceRefresh Then
         Call cleanDestRoot (false)
         Call renameFullyUsedFiles (false)
      End If

      Err.Description = "Unpack phase: destination directory"
      Err.Raise (up_ret)
   ElseIf up_ret = 900 Then
      If parm_bImportOnlyForceRefresh Then
         Call cleanDestRoot (false)
         Call renameFullyUsedFiles (false)
      End If

      Err.Description = "Unpack phase: Missing parameters"
      Err.Raise (up_ret)
   End If

   If parm_bImportOnlyForceRefresh Then Call cleanDestRoot (true)

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Sub


' #####
' ## cleanDestRoot (ByVal cdr_in_bIcsTmp)
' ## return: none
Sub cleanDestRoot (ByVal cdr_in_bIcsTmp)
   Dim fn: fn = "cleanDestRoot"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   Dim cdr_item

   For Each cdr_item In gl_aFullyUsedFiles
      If cdr_in_bIcsTmp Then cdr_item = cdr_item & cIcsTmpMark
      'Call gl_oFso.DeleteFile (up_fileName, True) -- Permission issue
      Call gl_oShell.Run ("CMD /C DEL /F /Q /A """& cdr_item &"""", 0 , True)
   Next

   For Each cdr_item In gl_aFullyUsedDirs
      If cdr_in_bIcsTmp Then cdr_item = cdr_item & cIcsTmpMark
      'Call gl_oFso.DeleteFolder (up_fileName, True) -- Permission issue
      Call gl_oShell.Run ("CMD /C RMDIR /S /Q """& cdr_item &"""", 0 , True)
   Next

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Sub


' #####
' ## renameFullyUsedFiles (ByVal rfur_in_bIcsTmp)
' ## return: none
Sub renameFullyUsedFiles (ByVal rfur_in_bIcsTmp)
   Dim fn: fn = "renameFullyUsedFiles"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   Dim rfuf_item, rfuf_file1, rfuf_file2

   For Each rfuf_item In gl_aFullyUsedFiles
      rfuf_file2 = gl_oFso.GetFileName (rfuf_item)

      If rfur_in_bIcsTmp Then
         rfuf_file1 = rfuf_item
         rfuf_file2 = rfuf_file2 & cIcsTmpMark
      Else
         rfuf_file1 = rfuf_item & cIcsTmpMark
         rfuf_file2 = rfuf_file2
      End If

      Call gl_oShell.Run ("CMD /C REN """& rfuf_file1 &""" """& rfuf_file2 &"""", 0 , True)
   Next

   For Each rfuf_item In gl_aFullyUsedDirs
      rfuf_file2 = gl_oFso.GetFileName (rfuf_item)

      If rfur_in_bIcsTmp Then
         rfuf_file1 = rfuf_item
         rfuf_file2 = rfuf_file2 & cIcsTmpMark
      Else
         rfuf_file1 = rfuf_item & cIcsTmpMark
         rfuf_file2 = rfuf_file2
      End If

      Call gl_oShell.Run ("CMD /C REN """& rfuf_file1 &""" """& rfuf_file2 &"""", 0 , True)
   Next

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Sub


' #####
' ## ics_rename (ByVal in_r_fileOrigName)
' ## return: none
Sub ics_rename (ByVal in_r_fileOrigName)
   Dim fn: fn = "ics_rename"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   Dim r_parms, r_cmd, r_ret, r_linkRemotePath

   If gl_linkDirName <> "" Then 
      r_linkRemotePath = gl_linkRemoteRoot &"/"& gl_linkDirName

      r_parms = parm_ftp &" "& parm_user &" "& parm_pass &" "& in_r_fileOrigName &" "& gl_oSupport.getMyTimeStamp () &" "& r_linkRemotePath &" """& gl_scriptWorkDir &""""
      r_cmd = """"& gl_appLoc &"\ics_rename.bat"" "& r_parms
      r_ret = gl_oShell.Run (r_cmd &" >> """& gl_scriptOutLoc &"""", 0, true)
   End If

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Sub


' #####
' ## putFile (ByVal in_pf_localName, ByVal in_pf_remoteName, ByVal in_pf_linkDirName)
' ## return: true || false
Function putFile (ByVal in_pf_localName, ByVal in_pf_remoteName, ByVal in_pf_linkDirName) ' return: true || false
   Dim fn: fn = "putFile"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   Dim pf_linkRemotePath: pf_linkRemotePath = gl_linkRemoteRoot
   If in_pf_linkDirName <> "" Then  pf_linkRemotePath = pf_linkRemotePath &"/"& in_pf_linkDirName
   If in_pf_remoteName = "" Then in_pf_remoteName = in_pf_localName ' !!!

   putFile = False

   Call muCreateMkdirCmdFile (pf_linkRemotePath, gl_mkdirCmdFileLoc)

   Dim pf_ret: pf_ret = ics_put (in_pf_localName, in_pf_remoteName, pf_linkRemotePath)

   If pf_ret = 500 Then
      Err.Description = " # "& fn &": Transfer Failed"
      Err.Raise (pf_ret)
   ElseIf pf_ret = 510 Then
      Err.Description = " # "& fn &": Not enough space"
      Err.Raise (pf_ret)
   End If

   Dim pf_aFileList: pf_aFileList = Array ()
   Dim pf_bFileExist: pf_bFileExist = False
   Dim pf_fileCnt, pf_aLine, pf_line, pf_size

   pf_ret = gl_oSupport.readFileToArray (gl_oFso, gl_oFso.BuildPath (gl_scriptWorkDir, "ics_confirm.list"), pf_aFileList, pf_fileCnt)
   If pf_ret <> "" Then
      Err.Description = " # "& fn &": Failed to confirm transfer for "& in_pf_remoteName &" -- "& pf_ret
      Err.Raise (19520)
   End If

   If pf_fileCnt < 3 Then
      Err.Description = " # "& fn &": Failed to transfer "& in_pf_remoteName &", empty directory"
      Err.Raise (19521)
   End If

   'If InStr (in_pf_localName, gl_scriptWorkDir) <> 1 Then in_pf_localName = gl_oFso.BuildPath (gl_scriptWorkDir, in_pf_localName)
   If InStr (in_pf_localName, ":") <> 2 Then in_pf_localName = gl_oFso.BuildPath (gl_scriptWorkDir, in_pf_localName) ' because in file mode in_pf_localName contains full path
   Dim pf_oFile: Set pf_oFile = gl_oFso.GetFile (in_pf_localName)

   Dim pf_oRegEx: Set pf_oRegEx = CreateObject("VBScript.RegExp")
   pf_oRegEx.IgnoreCase = True
   pf_oRegEx.Pattern = "\d\s+"& gl_oStringManip.escapeString (in_pf_remoteName) &"$"

   For Each pf_line In pf_aFileList
      If pf_oRegEx.Test (pf_line) Then
         pf_ret = gl_oStringManip.replaceRegExp (pf_line, "\s+", " ")
         pf_aLine = Split (pf_ret, " ")
         pf_size = Int (pf_aLine(4))

         If (pf_size = 0) Then
            Err.Description = " # "& fn &": Failed to transfer "& in_pf_remoteName &", ZERO size"
            Err.Raise (19523)
         Else
            pf_bFileExist = True
            If pf_oFile.Size <> pf_size Then Call log2file ("Transfer phase: file sizes are different for "& in_pf_remoteName &". local size: "& pf_oFile.Size &", remote size: "& pf_size, "W")
         End If
      End If
   Next

   If pf_bFileExist = False Then
      Err.Description = " # "& fn &": Failed to transfer "& in_pf_remoteName &", missing file"
      Err.Raise (19524)
   End If

   putFile = True

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)

   Set pf_oFile = Nothing
   Set pf_oRegEx = Nothing
End Function


' #####
' ## ics_put (ByVal in_p_localName, ByVal in_p_remoteName, ByVal in_p_remotePath)
' ## return: errorCode
Function ics_put (ByVal in_p_localName, ByVal in_p_remoteName, ByVal in_p_remotePath) ' return: errorCode
   Dim fn: fn = "ics_put"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   Dim p_parms, p_cmd

   ics_put = 999

   p_parms = parm_ftp &" "& parm_user &" "& parm_pass &" """& in_p_localName &""" """& in_p_remotePath &""" """& in_p_remoteName &""" """& gl_scriptWorkDir &""""
   p_cmd = """"& gl_appLoc &"\ics_put.bat"" "& p_parms
   ics_put = gl_oShell.Run (p_cmd &" >> """& gl_scriptOutLoc &"""", 0, true)

   ' copy output form ics_put to main out file
   p_parms = """"& gl_scriptOutLoc &""" """& gl_scriptWorkDir &"\ics_put.out"""
   p_cmd = """"& gl_appLoc &"\ics_merge.bat"" "& p_parms
   Call gl_oShell.Run (p_cmd, 0, true)

   p_parms = parm_ftp &" "& parm_user &" "& parm_pass &" """& in_p_remotePath &""" """& gl_scriptWorkDir &""""
   p_cmd = """"& gl_appLoc &"\ics_confirm.bat"" "& p_parms
   ics_put = gl_oShell.Run (p_cmd &" >> """& gl_scriptOutLoc &"""", 0, true)

   ' copy output form ics_confirm to main out file
   p_parms = """"& gl_scriptOutLoc &""" """& gl_scriptWorkDir &"\ics_confirm.out"""
   p_cmd = """"& gl_appLoc &"\ics_merge.bat"" "& p_parms
   Call gl_oShell.Run (p_cmd, 0, true)

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Function


' #####
' ## pack ()
Sub pack ()
   Dim fn: fn = "pack"
   Dim ret

   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   If parm_linkType = gl_icsLinkType.file And parm_bPackFileImportExport Then
      Dim p_item

      For Each p_item In parm_oFilesToProcess.Keys
         If parm_oFilesToProcess(p_item).exportRequest = True Then
            If Not ics_pack (p_item, p_item & gl_toPackExt) Then
               Err.Description =  "Failed to pack "& parm_linkType &" or it could be in use: "& p_item
               Err.Raise (19401)
             End If
          End If
      Next
   ElseIf parm_linkType = gl_icsLinkType.directory Then
      If Not ics_pack ("", gl_linkExportFile) Then
         Err.Description =  "Failed to pack "& parm_linkType &" or it could be in use."
         Err.Raise (19400)
      End If
   End If

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Sub


' #####
' ## ics_pack (ByVal in_p_itemToPack, ByVal in_p_linkExportFile)
' ## return: true || false
Function ics_pack (ByVal in_p_itemToPack, ByVal in_p_linkExportFile) ' return: true || false
   Dim p_parms, p_cmd, p_ret

   Dim fn: fn = "ics_pack"
   Dim p_noRecursion: p_noRecursion = ""
   Dim p_archType: p_archType = parm_archType

   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   If parm_bNoTopLevelFiles Or parm_bNoTopLevelFolders Then p_noRecursion = "norecursion"
   If parm_bNoTopLevelFolders Then in_p_itemToPack = "*.*"
   If parm_archType = "" Then p_archType = """"""

   ics_pack = True
   
   p_parms = """"& gl_syncRootDirExpand  &""" """& in_p_linkExportFile &""" "&_
            """"& gl_appLoc &""" """& in_p_itemToPack &""" "& parm_syncPass &" """& gl_scriptWorkDir &""" "& p_archType &" "& p_noRecursion
   p_cmd = """"& gl_appLoc &"\ics_pack.bat"" "& p_parms
   p_ret = gl_oShell.Run (p_cmd &" >> """& gl_scriptOutLoc &"""", 0, true)

   ' copy output form ics_pack to main out file
   p_parms = """"& gl_scriptOutLoc &""" """& gl_scriptWorkDir &"\ics_pack.out"""
   p_cmd = """"& gl_appLoc &"\ics_merge.bat"" "& p_parms
   Call gl_oShell.Run (p_cmd, 0, true)

   If p_ret = 400 Then
      Call log2file ("Pack phase: Packing problem for "& in_p_linkExportFile, "E")
      ics_pack = False
   ElseIf p_ret = 900 Then
      Call log2file ("Pack phase: Missing parameters", "W")
      ics_pack = False
   End If

   ' copy ics_pack.warr to log file
   Dim p_oWarrFile: Set p_oWarrFile = gl_oFso.GetFile (gl_scriptWorkDir &"\ics_pack.warr")
   If p_oWarrFile.Size > 0 Then
      p_parms = """"& gl_appSetupFilePath &"\"& cAppLogFile &""" """& gl_scriptWorkDir &"\ics_pack.warr"""
      p_cmd = """"& gl_appLoc &"\ics_merge.bat"" "& p_parms
      Call gl_oShell.Run (p_cmd, 0, true)
   End If
   Set p_oWarrFile = Nothing

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Function


' #####
' ## moveToBackUp ()
Sub moveToBackUp ()
   Dim mtbu_itemToDel, mtbu_item

   Dim fn: fn = "moveToBackUp"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   If parm_bBackUpRunImport Or parm_bBackUpRunExport Or parm_syncType = gl_icsSyncType.backupOnly Then
      Call gl_oSupport.makeDir (gl_oFso, parm_backUpDir)

      If parm_linkType = gl_icsLinkType.file Then
         Dim state, file, reFileName

         For Each mtbu_item In parm_oFilesToProcess.Keys
            file = mtbu_item
            state = mtbu_item & cStateExt

            If parm_bPackFileImportExport Then file = file & gl_toPackExt

            reFileName = gl_oStringManip.escapeString (file)

            mtbu_itemToDel = muGetOldestBackUp (parm_backUpDir, "^bakup\d{12}_"& reFileName &"$", parm_backUpCount)

            If mtbu_itemToDel <> "" Then Call gl_oFso.DeleteFile (parm_backUpDir &"\"& mtbu_itemToDel, True)

            If parm_bPackFileImportExport Then 
               Call gl_oFso.MoveFile (gl_scriptWorkDir &"\"& file, parm_backUpDir &"\bakup"& gl_oSupport.getMyTimeStamp() &"_"& file)
            Else 
               Call gl_oFso.CopyFile (gl_syncRootDirExpand &"\"& file, parm_backUpDir &"\bakup"& gl_oSupport.getMyTimeStamp() &"_"& file)
            End If

            If gl_oFso.FileExists (parm_backUpDir &"\"& state) Then Call gl_oFso.DeleteFile (parm_backUpDir &"\"& state, True)

            Call gl_oFso.MoveFile (gl_scriptWorkDir &"\"& state & cfnCurrSuff, parm_backUpDir &"\"& state)
         Next
      Else 
         mtbu_itemToDel = muGetOldestBackUp (parm_backUpDir, "^bakup\d{12}_"& gl_linkExportFile &"$", parm_backUpCount)
         If mtbu_itemToDel <> "" Then Call gl_oFso.DeleteFile (parm_backUpDir &"\"& mtbu_itemToDel, True)

         gl_oFso.MoveFile gl_linkExportFileLoc, parm_backUpDir &"\bakup"& gl_oSupport.getMyTimeStamp() &"_"& gl_linkExportFile

         If gl_oFso.FileExists (parm_backUpDir &"\"& gl_stateFile) Then Call gl_oFso.DeleteFile (parm_backUpDir &"\"& gl_stateFile, True)

         gl_oFso.MoveFile gl_stateFileCurrLoc, parm_backUpDir &"\"& gl_stateFile
      End If
   End If

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Sub


' #####
' ## runJob (ByVal in_rj_jobType, ByVal in_rj_flag)
' ##
' ## in_rj_jobType -- import || export
' ## in_rj_flag -- RunBefore || RunAfter
Sub runJob (ByVal in_rj_jobType, ByVal in_rj_flag)
   Dim rj_i, rj_hJobs, rj_jobName

   Dim fn: fn = "runJob"& in_rj_jobType
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   If in_rj_jobType = "export" Then
      If in_rj_flag = "RunBefore" Then
         Set rj_hJobs = parm_hExportJobToRunBefore
         rj_jobName = "exportJobToRunBefore"
      Else
         Set rj_hJobs = parm_hExportJobToRunAfter
         rj_jobName = "exportJobToRunAfter"
      End If
   Else
      in_rj_jobType = "import"

      If in_rj_flag = "RunBefore" Then
         Set rj_hJobs = parm_hImportJobToRunBefore
         rj_jobName = "importJobToRunBefore"
      Else
         Set rj_hJobs = parm_hImportJobToRunAfter
         rj_jobName = "importJobToRunAfter"
      End If
   End If

   For Each rj_i In rj_hJobs.Keys
      On Error Resume Next

      Call gl_oShell.Run (rj_i, 1, true)

      If Err Then Call log2file (" # runJob: failed to run "& in_rj_jobType &" "& rj_jobName &" ["& rj_i &"] -- "& Err.Number &":"& Err.Description, "E")

      On Error GoTo 0
   Next

   Set rj_hJobs = Nothing

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Sub


' #####
' ## updateConfigFile ()
Sub updateConfigFile ()
   Dim fn: fn = "updateConfigFile"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   Dim ucf_users, ucf_userId, ucf_configId

   Call gl_oSupport.getCurrUserId (ucf_users)
   ucf_userId = UCase (Join (ucf_users.Keys, ""))
   ucf_configId = gl_oSupport.getCompName() &"_"& ucf_userId

   Dim ucf_oXmlConfig, ucf_oConfigRoot, ucf_oMyConfig, ucf_oConfigNode, ucf_ret
   Dim ucf_remotePath: ucf_remotePath = ""
   Dim ucf_fnLinkConfigData: ucf_fnLinkConfigData = gl_linkDirName & cConfigFileExt 
   
   Set ucf_oXmlConfig = WScript.CreateObject ("Microsoft.XMLDOM")
   ucf_oXmlConfig.async = "false"

   If parm_syncType <> gl_icsSyncType.exportOnly Then
      ucf_remotePath = gl_linkDirName

      ucf_ret = getFile (ucf_fnLinkConfigData, ucf_fnLinkConfigData, ucf_remotePath)
      If ucf_ret Then
         ucf_oXmlConfig.Load (gl_oFso.BuildPath (gl_scriptWorkDir, ucf_fnLinkConfigData))

         If ucf_oXmlConfig.parseError.errorCode Then
            Set ucf_oConfigRoot = ucf_oXmlConfig.createElement ("iCanSyncConfig")
            ucf_oXmlConfig.documentElement = ucf_oConfigRoot
         Else
            Set ucf_oConfigRoot = ucf_oXmlConfig.selectSingleNode ("//iCanSyncConfig")         
         End If

         ' to remove old configuration
         Set ucf_oConfigNode = ucf_oXmlConfig.selectSingleNode ("//"& ucf_configId)
         If Not ucf_oConfigNode Is Nothing Then ucf_oConfigNode.parentNode.removeChild (ucf_oConfigNode)

         ' add new configuration
         Set ucf_oMyConfig = ucf_oXmlConfig.createElement (ucf_configId)
         ucf_oConfigRoot.appendChild (ucf_oMyConfig)
      End If
   Else
      Set ucf_oConfigRoot = ucf_oXmlConfig.createElement ("iCanSyncConfig")
      ucf_oXmlConfig.documentElement = ucf_oConfigRoot

      Set ucf_oMyConfig = ucf_oXmlConfig.firstChild
   End If

   On Error Resume Next

   Set ucf_oConfigNode = oNode.selectSingleNode ("./linkName")
   ucf_oMyConfig.appendChild (ucf_oConfigNode)

   Set ucf_oConfigNode = oNode.selectSingleNode ("./linkType")
   ucf_oMyConfig.appendChild (ucf_oConfigNode)

   Set ucf_oConfigNode = oNode.selectSingleNode ("./syncType")
   ucf_oMyConfig.appendChild (ucf_oConfigNode)

   Set ucf_oConfigNode = oNode.selectSingleNode ("./syncRootDir")
   ucf_oMyConfig.appendChild (ucf_oConfigNode)

   Set ucf_oConfigNode = oNode.selectSingleNode ("./syncWorkDir")
   ucf_oMyConfig.appendChild (ucf_oConfigNode)

   Set ucf_oConfigNode = oNode.selectSingleNode ("./packFormat")
   ucf_oMyConfig.appendChild (ucf_oConfigNode)

   Set ucf_oConfigNode = oNode.selectSingleNode ("./packFileImportExport")
   ucf_oMyConfig.appendChild (ucf_oConfigNode)

   Set ucf_oConfigNode = oNode.selectSingleNode ("./export")
   ucf_oMyConfig.appendChild (ucf_oConfigNode)

   Set ucf_oConfigNode = oNode.selectSingleNode ("./import")
   ucf_oMyConfig.appendChild (ucf_oConfigNode)

   Set ucf_oConfigNode = ucf_oXmlConfig.createElement ("hostName")
   ucf_oMyConfig.appendChild (ucf_oConfigNode)
   ucf_oConfigNode.Text = gl_oSupport.getCompName()

   Set ucf_oConfigNode = ucf_oXmlConfig.createElement ("hostUserId")
   ucf_oMyConfig.appendChild (ucf_oConfigNode)
   ucf_oConfigNode.Text = ucf_userId

   Set ucf_oConfigNode = ucf_oXmlConfig.createElement ("iter")
   ucf_oMyConfig.appendChild (ucf_oConfigNode)

   If oNode.selectSingleNode("./iter") Is Nothing Then 
      If Len (parm_syncPass) > 0 Then ucf_oConfigNode.Text = "1."& Len (parm_syncPass)
   Else
      ucf_oConfigNode.Text = oNode.selectSingleNode("./iter").Text
   End If

   ucf_oXmlConfig.save (gl_oFso.BuildPath (gl_scriptWorkDir, ucf_fnLinkConfigData))
   
   On Error Goto 0

   ucf_ret = putFile (ucf_fnLinkConfigData, ucf_fnLinkConfigData, ucf_remotePath)

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Sub


' #####
' ## Class::CIcsFileInfo 
Class CIcsFileInfo
   Dim fileName
   Dim isRunNeeded ' if 0 then no need; 1 - run export; 2 - run all;
   Dim importRequest
   Dim exportRequest
End Class

' Set oIcsFileInfo = New CIcsFileInfo
' Set oExportFileInfo = CreateObject("Scripting.Dictionary")
' Set oFileInfo = New CIcsFileInfo
' oExportFileInfo.fileName = 
' oExportFileInfo.isRunNeeded = 
' oExportFileInfo.Add fileName, oExportFileInfo

' For Each key In oExportFileInfo.Keys
'    oExportFileInfo(key).fileName
' Next


' #####
' ## muGetOldestBackUp (ByVal in_gobu_where, ByVal in_gobu_reWhat, ByVal in_gobu_minCount)
' ## return: oldest file name
Function muGetOldestBackUp (ByVal in_gobu_where, ByVal in_gobu_reWhat, ByVal in_gobu_minCount) ' return: oldest file name
   Dim gobu_oFolder, gobu_oFiles, gobu_oFile
   Dim gobu_oRegEx
   Dim gobu_cntBackup, gobu_oldestBackUp
   Dim gobu_fileName
   Dim gobu_colMatches

   Dim fn: fn = "muGetOldestBackUp:Fatal error to find oldest backup file at """& in_gobu_where &""""
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   Set gobu_oFolder = gl_oFso.GetFolder (in_gobu_where)
   Set gobu_oFiles = gobu_oFolder.Files

   Set gobu_oRegEx = CreateObject("VBScript.RegExp")
   gobu_oRegEx.IgnoreCase = True
   gobu_oRegEx.Pattern = in_gobu_reWhat

   gobu_cntBackup = 0
   gobu_oldestBackUp = ""

   For Each gobu_oFile In gobu_oFiles
      gobu_fileName = LCase (gobu_oFile.Name)

      Set gobu_colMatches = gobu_oRegEx.Execute (gobu_fileName)

      If gobu_colMatches.Count > 0 Then
         gobu_cntBackup = gobu_cntBackup + 1

         If StrComp (gobu_oldestBackUp, gobu_fileName) > 0 Or gobu_oldestBackUp = "" Then
            gobu_oldestBackUp = gobu_fileName
         End If
      End If
   Next
   
   Set gobu_oFolder = Nothing
   Set gobu_oFiles = Nothing
   Set gobu_oFile = Nothing
   Set gobu_oRegEx = Nothing
   Set gobu_colMatches = Nothing

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)

   If gobu_cntBackup > in_gobu_minCount - 1 Then muGetOldestBackUp = gobu_oldestBackUp
End Function


' #####
' ## muCreateMkdirCmdFile (ByVal in_cmcf_path, ByVal in_cmcf_mkdirCmdFile)
Sub muCreateMkdirCmdFile (ByVal in_cmcf_path, ByVal in_cmcf_mkdirCmdFile)
   Dim cmcf_aFiles, cmcf_pathPart
   Dim cmcf_oStream
   Dim cmcf_i

   Dim fn: fn = "muCreateMkdirCmdFile:Fatal error to create "& in_cmcf_mkdirCmdFile &" for path """& in_cmcf_path &""""
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   cmcf_aFiles = Split(in_cmcf_path, "/")
   cmcf_pathPart = ""

   Set cmcf_oStream = gl_oFso.CreateTextFile (in_cmcf_mkdirCmdFile, True, False) ' !!! non ASCII-7 characters are not supported in remote names

   For cmcf_i = 0 to Ubound(cmcf_aFiles)
      If cmcf_pathPart = "" Then
         cmcf_pathPart = cmcf_aFiles(cmcf_i)
      Else
         cmcf_pathPart = cmcf_pathPart &"/"& cmcf_aFiles(cmcf_i)
      End If

      cmcf_oStream.WriteLine "mkdir "& cmcf_pathPart
   Next

   cmcf_oStream.close

   Set cmcf_oStream = Nothing

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Sub


' #####
' ## muSaveFileFromWeb (ByVal in_sffw_url, ByVal in_sffw_file)
' ## return: true || false
Function muSaveFileFromWeb (ByVal in_sffw_url, ByVal in_sffw_file) ' return: true || false
   Const sffw_cTypeBinary = 1
   Const sffw_cTypeText = 2
   Const sffw_cSaveCreateNotExist  = 1
   Const sffw_cSaveCreateOverWrite = 2

   Dim sffw_oHttp, sffw_oAdo, sffw_aByteResponse

   Dim fn: fn = "muSaveFileFromWeb"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   muSaveFileFromWeb = False

   On Error Resume Next

   'Call log2file (" # "& fn &": save ["& in_sffw_file &"]", "I") ' for debugging
   Call log2file (" # "& fn &": downloading ["& in_sffw_url &"]", "I")

   'Set sffw_oHttp = CreateObject("Microsoft.XMLHTTP")
   'If WScript.Version = "5.6" Then
      Set sffw_oHttp = CreateObject("MSXML2.XMLHTTP")
   'Else
      'Set sffw_oHttp = CreateObject("WinHttp.WinHttpRequest.5.1") ' !!! doesn't work with 2k
   'End If

   in_sffw_url = Replace (in_sffw_url, "//", "/")
   in_sffw_url = Replace (in_sffw_url, ":/", "://")

   sffw_oHttp.Open "GET", in_sffw_url, false 
   sffw_oHttp.Send () ' If remote file doesn't exist then should fail here and NO local file will be created. !!!

   If Err Then
      Set sffw_oHttp = Nothing
      Dim errNum, errDesc

      errNum = Err.Number
      errDesc = Err.Description

      If errNum = -2147024891 Or errNum = -2147012867 Then
         Call log2file (" # "& fn &": "& errDesc &" ["& in_sffw_url &"]", "W")
      Else 
         On Error Goto 0
         Err.Description = errDesc &" -- "& in_sffw_url
         Err.Raise (errNum)
      End If

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
      Exit Function
   End If

   If sffw_oHttp.Status <> "200" Then
      Set sffw_oHttp = Nothing

      On Error Goto 0

      Err.Description = "Not a successful response for requested URL -- "& in_sffw_url
      Err.Raise (19014)

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
      Exit Function
   End If

   sffw_aByteResponse = sffw_oHttp.ResponseBody
   Set sffw_oHttp = Nothing

   Set sffw_oAdo = Nothing
   Set sffw_oAdo = CreateObject("ADODB.Stream") 
   If sffw_oAdo Is Nothing Then
      Call log2file (" # "& fn &": No ADODB", "I")

'NIU       Set oFile = gl_oFso.OpenTextFile(in_sffw_file, sffw_cSaveCreateOverWrite, True)

'NIU       For cnt = 0 to UBound(sffw_aByteResponse)
'NIU          oFile.Write Chr(255 And Ascb(Midb(sffw_aByteResponse,cnt + 1, 1)))
'NIU       Next
'NIU       oFile.Close

'NIU       Set oFile = Nothing
   Else
      Call log2file (" # "& fn &": ADODB in use", "I")

      sffw_oAdo.Open
      sffw_oAdo.Type = sffw_cTypeBinary
      sffw_oAdo.Write sffw_aByteResponse
      sffw_oAdo.Position = 0
      sffw_oAdo.SaveToFile in_sffw_file, sffw_cSaveCreateOverWrite 
      sffw_oAdo.Close 

      muSaveFileFromWeb = True
   End If

   Call log2file (" # "& fn &": done", "I")

   Set sffw_oAdo = Nothing
   Set sffw_aByteResponse = Nothing

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Function


' #####
' ## muExpandEnvironmentStrings (ByVal in_ees_str, ByRef in_ref_ees_expansionFlag)
' ## return: expanded string
' ## in_ref_ees_expansionFlag: "" - ok || error string;
Function muExpandEnvironmentStrings (ByVal in_ees_str, ByRef in_ref_ees_expansionFlag) ' return: expanded string
'   Dim ees_envValue, ees_envKey, ees_colUsers, ees_userId
'   Dim ees_procEnv, ees_aStrSplit
'   Dim ees_sysUserProfile, ees_userProfilesPath, ees_UserProfile

   Dim fn: fn = "muExpandEnvironmentStrings"
   If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

   in_ref_ees_expansionFlag = ""

   ' currUserId
'   ees_envValue = gl_oSupport.getCurrUserId (ees_colUsers)
'
'   If ees_envValue = 0 Or ees_envValue = 1 Then
'      ees_userId = Join (ees_colUsers.Keys, "")
'   Else
'      ees_userId = ""
'      in_ref_ees_expansionFlag = "NoUserId"
'   End If

   ' looking for %UserProfile%
'   Set ees_procEnv = gl_oShell.Environment ("Process")
'   ees_sysUserProfile = gl_oStringManip.cleanPathString (ees_procEnv.Item("USERPROFILE"))
'   ees_userProfilesPath = Left (ees_sysUserProfile,  InstrRev (ees_sysUserProfile, "\"))
'   ees_UserProfile = ees_userProfilesPath & ees_userId

   ' ExpandEnvironmentStrings
'   ees_aStrSplit = Split (in_ees_str, "%")
'   For Each ees_envKey In ees_aStrSplit
'      If ees_envKey <> "" Then
'         ees_envValue = gl_oShell.SpecialFolders(ees_envKey)
'         if ees_envValue <> "" Then in_ees_str = Replace (in_ees_str, "%"& ees_envKey &"%", ees_envValue)
'      End If
'   Next
'
'   If InStr (in_ees_str, "%") Then
'      in_ees_str = gl_oShell.ExpandEnvironmentStrings (in_ees_str)
'   End If

'   muExpandEnvironmentStrings = Replace (in_ees_str, ees_sysUserProfile, ees_UserProfile)

   muExpandEnvironmentStrings = gl_oShell.ExpandEnvironmentStrings (in_ees_str)

   If InStr (muExpandEnvironmentStrings, "%") Then
      in_ref_ees_expansionFlag = "NoUserId or couldn't expand environment variables"
   End If

   muExpandEnvironmentStrings = gl_oStringManip.cleanPathString (muExpandEnvironmentStrings)

'   Set ees_aStrSplit = Nothing
'   Set ees_procEnv = Nothing

   If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
End Function


' #####
' ## muChangeFileEncoding (ByVal in_cfe_fileIn, ByVal in_cfe_fileOut, ByVal in_cfe_encoding)
' ## return: true || false
' ##
' ## "us-ascii", "Unicode", "utf-8", "koi8-r"
Function muChangeFileEncoding (ByVal in_cfe_fileIn, ByVal in_cfe_fileOut, ByVal in_cfe_encoding) ' return: true || false
   Const cfe_cTypeBinary = 1
   Const cfe_cTypeText = 2
   Const cfe_cSaveCreateNotExist  = 1
   Const cfe_cSaveCreateOverWrite = 2

   On Error Resume Next
   Dim cfe_oAdoStrmIn: Set cfe_oAdoStrmIn = CreateObject("ADODB.Stream") 
   Dim cfe_oAdoStrmOut: Set cfe_oAdoStrmOut = CreateObject("ADODB.Stream") 

   cfe_oAdoStrmIn.Open
   cfe_oAdoStrmIn.Type = cfe_cTypeText
   cfe_oAdoStrmIn.Position = 0
'   cfe_oAdoStrmIn.Charset = "Unicode"
   cfe_oAdoStrmIn.LoadFromFile in_cfe_fileIn

   cfe_oAdoStrmOut.Open
   cfe_oAdoStrmOut.Type = cfe_cTypeText
   cfe_oAdoStrmOut.Charset = in_cfe_encoding
   cfe_oAdoStrmOut.WriteText cfe_oAdoStrmIn.ReadText
   cfe_oAdoStrmOut.SaveToFile in_cfe_fileOut, cfe_cSaveCreateOverWrite 

   cfe_oAdoStrmIn.Close
   cfe_oAdoStrmOut.Close

   Set cfe_oAdoStrmIn  = Nothing
   Set cfe_oAdoStrmOut  = Nothing

   If Err Then
      muChangeFileEncoding = False
   Else
      muChangeFileEncoding = True
   End If
    
   On Error Goto 0
End Function


'Call log2file (Err.Number & ": " & Err.Description &" -- "& gl_scriptErrBuff, "E")
'Err.Clear

'ADODB.Stream error '800a0bbc' 


' #####
' ## includes
' #####

'Option Explicit
'
'ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile ("TCPconsulting_debug.vbs").ReadAll
'ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile ("TCPconsulting_CStringManip.vbs").ReadAll
'ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile ("TCPconsulting_CSupport.vbs").ReadAll
'
'Dim obj: Set obj = new CFileStructure
'
'Dim aExcludedFile: aExcludedFile = Array ()
'Dim aExcludedPath: aExcludedPath = Array ("*_Debug\")
'
'MsgBox obj.ClassName
'MsgBox obj.setStartingPoint ("C:\Trash\build_iCanSync\")
'
'Call obj.populateStructure ("", aExcludedPath, aExcludedFile, "")
'Call obj.saveStructure ("C:\Trash\out.xml")
'MsgBox obj.getStructure().xml
'
'MsgBox "workItem -- file"
'aExcludedPath = Array ("C:\Trash\build_iCanSync\testSetupFiles", "C:\Trash\build_iCanSync\oldVersions")
'Call obj.populateStructure ("ToDo.txt", aExcludedPath, aExcludedFile, "")
'Call obj.saveStructure ("C:\Trash\out_fileOnly.xml")
'
'MsgBox "workItem -- Folder & no sub folders"
'aExcludedPath = Array ("$NOTOPLEVELFOLDERS$")
'aExcludedFile = Array ()
'Call obj.populateStructure ("", aExcludedPath, aExcludedFile, "FullPath|AllDates")
'Call obj.saveStructure ("C:\Trash\out_workdir_noSubFolders.xml")
'
'MsgBox "workItem -- Folder & no top files"
'aExcludedPath = Array ()
'aExcludedFile = Array ("$NOTOPLEVELFILES$", "*.dll", "*.TXT")
'Call obj.populateStructure ("", aExcludedPath, aExcludedFile, "AllDates")
'Call obj.saveStructure ("C:\Trash\out_workdir_noTopFiles.xml")
'
'MsgBox "workItem -- Folder"
'aExcludedPath = Array ("_Release\iCanSyncEZ\7zip")
'aExcludedFile = Array ()
'Call obj.populateStructure ("", aExcludedPath, aExcludedFile, "FullPath|AllDates")
'
'Dim oStream
'Dim oFso: Set oFso = WScript.CreateObject ("Scripting.FileSystemObject")
'Set oStream = oFso.OpenTextFile ("C:\Trash\FileStruct_out.txt", 2, True, -1)
'oStream.Write Join (obj.getFullyUsedFolders (), vbcrlf)
'oStream.Write vbcrlf
'oStream.Write Join (obj.getFullyUsedFiles (), vbcrlf)
'oStream.Close
'Set oStream = Nothing

'Call obj.saveStructure ("C:\Trash\out_workdir.xml")
'
'MsgBox "workItem -- WorkingDir"
'aExcludedPath = Array ()
'aExcludedFile = Array ("ToDo.txt", "7z.exe")
''Call obj.populateStructure ("iCanSyncEZ", aExcludedPath, aExcludedFile, "FullPath|AllDates")
'Call obj.populateStructure ("", aExcludedPath, aExcludedFile, "FullPath|AllDates")
'Call obj.saveStructure ("C:\Trash\out_workdir_workingDir.xml")

'MsgBox obj.getFileById (obj.generateFileId("C:\Trash\build_iCanSync\iCanSyncEZ\7zip\7z.exe")).xml
'
'Dim cmpRet: Set cmpRet = obj.compareStructure ("C:\Trash\out_workdir.xml")
'
'MsgBox cmpRet("status")
'MsgBox "Desc: "& cmpRet("description")
'
'MsgBox "New files: "&vbcrlf& Join (cmpRet("newFiles"), vbcrlf)
'Set oStream = oFso.OpenTextFile ("C:\Trash\out_newFiles.txt", 2, True, -1)
'oStream.Write Join (cmpRet("newFiles"), vbcrlf)
'oStream.Close
'
'MsgBox "Mod files: "&vbcrlf& Join (cmpRet("modifFiles"), vbcrlf)
'Set oStream = oFso.OpenTextFile ("C:\Trash\out_modifFiles.txt", 2, True, -1)
'oStream.Write Join (cmpRet("modifFiles"), vbcrlf)
'oStream.Close
'
'MsgBox "Del files: "&vbcrlf& Join (cmpRet("delFiles"), vbcrlf)
'Set oStream = oFso.OpenTextFile ("C:\Trash\out_delFiles.txt", 2, True, -1)
'oStream.Write Join (cmpRet("delFiles"), vbcrlf)
'oStream.Close
'
'MsgBox "Excl files: "&vbcrlf& Join (cmpRet("exclFiles"), vbcrlf)
'Set oStream = oFso.OpenTextFile ("C:\Trash\out_exclFiles.txt", 2, True, -1)
'oStream.Write Join (cmpRet("exclFiles"), vbcrlf)
'oStream.Close
'
'MsgBox "Path doesn't exist"
'MsgBox obj.setStartingPoint ("%TEMP%")
'Call obj.populateStructure ("ds", Array (), Array(), "")
'Call obj.saveStructure ("C:\Trash\out_nodir.xml")
'MsgBox "Done"


Const cNoTopLevelFolders = "$NOTOPLEVELFOLDERS$"
Const cNoTopLevelFiles = "$NOTOPLEVELFILES$"


Class CFileStructure ' !!! requires: CStringManip & CSupport
   Private m_className__

   Private m_path, m_oFolder
   Private m_oFso, m_oShell, m_oStringManip, m_oSupport
   Private m_oReWorkingDirSearch, m_oReExcludedSearch
   Private m_aExcludedPath, m_aExcludedFile
   Private m_oXmlFileStructure, m_oFileNodes
   Private m_oFileCount ' !!! real file count
   Private m_oFolderCount ' !!! folder count under StartingPoint minus excluded folders
   Private m_fFileInfo, m_fPathFormat
   Private m_fNoTopLevelFolders, m_fNoTopLevelFiles
   Private m_aFullyUsedDirs, m_aFullyUsedFiles


   Private Sub Class_Initialize
      m_className__ = "CFileStructure"

      Set m_oFso = CreateObject ("Scripting.FileSystemObject")
      Set m_oShell = CreateObject("WScript.Shell")
      Set m_oStringManip = New CStringManip
      Set m_oSupport = New CSupport
      Set m_oReWorkingDirSearch = New RegExp
      Set m_oReExcludedSearch = New RegExp
      Set m_oXmlFileStructure = Nothing
      Set m_oFolder = Nothing

      m_oReWorkingDirSearch.IgnoreCase = True
      m_oReExcludedSearch.IgnoreCase = True

      m_fNoTopLevelFolders = False
      m_fNoTopLevelFiles = False

      m_aFullyUsedDirs = Array ()
      m_aFullyUsedFiles = Array ()
   End Sub

   
   Public Property Get ClassName
      ClassName = m_className__
   End Property


   ' #####
   ' ## Class::CFileStructure::setStartingPoint (ByVal in_ssp_path)
   ' ## return: previous starting point
   Public Function setStartingPoint (ByVal in_ssp_path) ' return: previous starting point
      in_ssp_path = m_oShell.ExpandEnvironmentStrings (in_ssp_path)

      If Not InStr (in_ssp_path, ":\") = 2 Then
         Err.Description = "Valid path is required."
         Err.Raise (19001)
      End If

      setStartingPoint = m_path
      m_path = m_oStringManip.cleanPathString (in_ssp_path)

      resetStructure ()

      If m_oFso.FolderExists (in_ssp_path) Then 
         Set m_oFolder = m_oFso.GetFolder (in_ssp_path)
      Else
         Set m_oFolder = Nothing
      End If
   End Function


   ' #####
   ' ## Class::CFileStructure::populateStructure (ByVal in_ps_workItem, ByVal in_ps_aExcludedPath, ByVal in_ps_aExcludedFile, ByVal in_ps_Flags [flag1|flag2])
   ' ## return: file count
   ' ##
   ' ## !!! Help !!! 
   ' ## in_ps_workItem doesn't support wildcard
   Public Function populateStructure (ByVal in_ps_workItem, ByVal in_ps_aExcludedPath, ByVal in_ps_aExcludedFile, ByVal in_ps_Flags) ' return: file count
      If m_oFolder Is Nothing Then Exit Function

      resetStructure ()

      Dim aFlags: aFlags = Split (in_ps_Flags, "|")
      Dim item

      For Each item In aFlags
         If item = "AllDates" Then 
            m_fFileInfo = item
         ElseIf item = "FullPath" Then
            m_fPathFormat = item
         End If
      Next

      in_ps_workItem = m_oStringManip.cleanPathString (in_ps_workItem)

      Dim itemLoc: itemLoc = m_oFso.BuildPath (m_path, in_ps_workItem)
      If m_oFso.FileExists (itemLoc) Then ' !!! if in_ps_workItem file and folder (impossible but) then folder will be ignored
         Dim oFile: Set oFile = m_oFso.GetFile (itemLoc)
         Call addFileRec (oFile)

         populateStructure = 1
         m_oFileCount.Text = populateStructure
         m_oFolderCount.Text = 0

         Exit Function
      End If

      With m_oReWorkingDirSearch
         .Global = True

         If in_ps_workItem = "" Then
            .Pattern = ""
         Else
            .Pattern = "\\"& in_ps_workItem &"\\"
         End If
      End With


      For Each item In in_ps_aExcludedPath
         If item = cNoTopLevelFolders Then m_fNoTopLevelFolders = True
      Next

      If m_fNoTopLevelFolders Then 
         m_aExcludedPath = Array ("*")
      Else
         m_aExcludedPath = in_ps_aExcludedPath
      End If
      

      For Each item In in_ps_aExcludedFile
         If item = cNoTopLevelFiles Then
            m_fNoTopLevelFiles = True
         Else
            Call m_oSupport.arrayPush (m_aExcludedFile, item)
         End If
      Next

      If getFolderStructure (m_oFolder) And UBound (m_aFullyUsedDirs) = -1 And UBound (m_aFullyUsedFiles) = -1 Then Call m_oSupport.arrayPush (m_aFullyUsedDirs, m_oFolder.Path)

      m_oFolderCount.Text = Int (m_oFolderCount.Text) - 1

      populateStructure = m_oFileCount.Text
   End Function


   ' #####
   ' ## Class::CFileStructure::getFolderStructure (ByVal in_gfs_oFolder)
   Private Function getFolderStructure (ByVal in_gfs_oFolder)
      Dim bFullyUsedItems: bFullyUsedItems = True
      Dim aFullyUsedFiles: aFullyUsedFiles = Array ()
      Dim aFullyUsedDirs: aFullyUsedDirs = Array ()
      Dim colMatches

      Dim isAlias: isAlias = in_gfs_oFolder.Attributes And 1024
      If isAlias = 1024 Then Exit Function

      Set colMatches = m_oReWorkingDirSearch.Execute (in_gfs_oFolder.Path &"\") ' !!! &"\" just for m_oReWorkingDirSearch.Pattern
      If colMatches.Count > 0 And Not m_fNoTopLevelFiles Then
         Dim colFiles: Set colFiles = in_gfs_oFolder.Files
         Dim oFile

         For Each oFile In colFiles
            If isExcludedFile (oFile.Path) Then
               bFullyUsedItems = False
            Else
               Call addFileRec (oFile)
               Call m_oSupport.arrayPush (aFullyUsedFiles, oFile.Path)
            End If
         Next
      End If

      m_fNoTopLevelFiles = False


      m_oFolderCount.Text = Int (m_oFolderCount.Text) + 1

      Dim colFolders: Set colFolders = in_gfs_oFolder.SubFolders
      Dim oSubFolder

      For Each oSubFolder In colFolders
         If isExcludedPath (oSubFolder.Path) Then
            bFullyUsedItems = False
         Else
            If getFolderStructure (oSubFolder) Then
               Call m_oSupport.arrayPush (aFullyUsedDirs, oSubFolder.Path)
            Else
               bFullyUsedItems = False
            End If
         End If
      Next

      If Not bFullyUsedItems Then
         Call m_oSupport.arrayPush (m_aFullyUsedFiles, aFullyUsedFiles)
         Call m_oSupport.arrayPush (m_aFullyUsedDirs, aFullyUsedDirs)
      End If

      getFolderStructure = bFullyUsedItems
   End Function


   ' #####
   ' ## Class::CFileStructure::isExcludedFile (ByVal in_ief_fileLoc)
   ' ## return: true || false
   Public Function isExcludedFile (ByVal in_ief_fileLoc)
      Dim excludedItem, colMatches

      isExcludedFile = False

      For Each excludedItem In m_aExcludedFile
         excludedItem = m_oShell.ExpandEnvironmentStrings (excludedItem)
         excludedItem = m_oStringManip.cleanPathString (excludedItem)
         excludedItem = m_oStringManip.escapeString (excludedItem)
         excludedItem = Replace (excludedItem, "*", ".*")

         m_oReExcludedSearch.Pattern = "\\"& excludedItem &"$"
         Set colMatches = m_oReExcludedSearch.Execute (in_ief_fileLoc)

         If colMatches.Count > 0 Then
            isExcludedFile = True

            Exit Function
         End If
      Next
   End Function


   ' #####
   ' ## Class::CFileStructure::isExcludedPath (ByVal in_iep_path) path or location
   ' ## return: true || false
   Public Function isExcludedPath (ByVal in_iep_path)
      Dim excludedItem, colMatches

      isExcludedPath = False

      For Each excludedItem In m_aExcludedPath
         excludedItem = m_oShell.ExpandEnvironmentStrings (excludedItem)
         excludedItem = m_oStringManip.cleanPathString (excludedItem)
         excludedItem = m_oStringManip.escapeString (excludedItem)
         excludedItem = Replace (excludedItem, "*", ".*")

         m_oReExcludedSearch.Pattern = "\\"& excludedItem &"$"
         Set colMatches = m_oReExcludedSearch.Execute (in_iep_path)

         If colMatches.Count > 0 Then ' for path
            isExcludedPath = True

            Exit Function
         Else ' for location
            If InStr (excludedItem, ":\") <> 2 And Left (excludedItem, 1) <> "\" Then excludedItem = "\\"& excludedItem &"\\"

            m_oReExcludedSearch.Pattern = excludedItem
            Set colMatches = m_oReExcludedSearch.Execute (in_iep_path)

            If colMatches.Count > 0 Then
               isExcludedPath = True

               Exit Function
            End If
         End If
      Next
   End Function


   ' #####
   ' ## Class::CFileStructure::addFileRec (ByRef in_ref_afr_oFile)
   ' ## return: none
   Private Sub addFileRec (ByRef in_ref_afr_oFile)
      Dim oFileNode, oFileLoc, oFileDate, oFileSize
      Dim fileLoc: fileLoc = in_ref_afr_oFile.Path

      If m_fPathFormat <> "FullPath" Then
         fileLoc = m_oStringManip.replaceRegExp (fileLoc, m_oStringManip.escapeString (m_path &"\"), "")
      End If

      Set oFileNode = m_oXmlFileStructure.createElement ("file")
      Call oFileNode.setAttribute ("ID", generateFileId(fileLoc))
      m_oFileNodes.appendChild (oFileNode)

      Set oFileLoc = m_oXmlFileStructure.createElement ("location") 
      oFileNode.appendChild (oFileLoc)
      oFileLoc.Text = fileLoc

      Set oFileDate = m_oXmlFileStructure.createElement ("modifiedDate") 
      oFileNode.appendChild (oFileDate)
      oFileDate.Text = convertDate (in_ref_afr_oFile.DateLastModified)

      Set oFileSize = m_oXmlFileStructure.createElement ("size") 
      oFileNode.appendChild (oFileSize)
      oFileSize.Text = in_ref_afr_oFile.Size

      If m_fFileInfo = "AllDates" Then
         Set oFileDate = m_oXmlFileStructure.createElement ("createdDate") 
         oFileNode.appendChild (oFileDate)
         oFileDate.Text = convertDate (in_ref_afr_oFile.DateCreated)

         Set oFileDate = m_oXmlFileStructure.createElement ("accessedDate") 
         oFileNode.appendChild (oFileDate)
         oFileDate.Text = convertDate (in_ref_afr_oFile.DateLastAccessed)
      End If

      m_oFileCount.Text = Int (m_oFileCount.Text) + 1
   End Sub


   ' #####
   ' ## Class::CFileStructure::resetStructure ()
   Private Sub resetStructure ()
      Set m_oXmlFileStructure = WScript.CreateObject ("Microsoft.XMLDOM")
      Set m_oFileNodes = m_oXmlFileStructure.createElement ("files")
      Set m_oFileCount = m_oXmlFileStructure.createElement ("fileCount")
      Set m_oFolderCount = m_oXmlFileStructure.createElement ("folderCount")

      m_oFileCount.Text = 0
      m_oFolderCount.Text = 0

      Dim oRoot: Set oRoot = m_oXmlFileStructure.createElement ("root")
      
      m_oXmlFileStructure.documentElement = oRoot
      oRoot.appendChild (m_oFileCount)
      oRoot.appendChild (m_oFolderCount)
      oRoot.appendChild (m_oFileNodes)

      m_fFileInfo = ""
      m_fPathFormat = ""

      m_fNoTopLevelFolders = False
      m_fNoTopLevelFiles = False

      m_aExcludedPath = Array ()
      m_aExcludedFile = Array ()

      m_aFullyUsedFiles = Array ()
      m_aFullyUsedDirs = Array ()
   End Sub
   

   ' #####
   ' ## Class::CFileStructure::convertDate (ByVal in_cd_date)
   ' ## return: date string
   Private Function convertDate (ByVal in_cd_date) ' return: date string
      convertDate = DatePart ("yyyy", in_cd_date) &"/"& DatePart ("m", in_cd_date) &"/"& DatePart ("d", in_cd_date) &" "& DatePart ("h", in_cd_date) &":"& DatePart ("n", in_cd_date)
   End Function


   ' #####
   ' ## Class::CFileStructure::getStructure ()
   ' ## return: oXmlFileStructure
   Public Function getStructure () ' return: oXmlFileStructure
      If m_oXmlFileStructure is Nothing Then
         Err.Description = "setStartingPoint/populateStructure needs to be executed first."
         Err.Raise (19002)
      End If

      Set getStructure = m_oXmlFileStructure
   End Function


   ' #####
   ' ## Class::CFileStructure::compareStructure (ByVal in_cs_fileLoc)
   ' ## return: hash of changes
   Public Function compareStructure (ByVal in_cs_fileLoc) ' return: hash of changes
      Dim hRet: Set hRet = CreateObject ("Scripting.Dictionary")
      Dim oXmlOldStructure: Set oXmlOldStructure = WScript.CreateObject ("Microsoft.XMLDOM")
      Dim aNewFiles: aNewFiles = Array()
      Dim aModifFiles: aModifFiles = Array()
      Dim aDelFiles: aDelFiles = Array()
      Dim aExcludedFiles: aExcludedFiles = Array()

      in_cs_fileLoc = m_oShell.ExpandEnvironmentStrings (in_cs_fileLoc)

      If Not m_oFso.FileExists (in_cs_fileLoc) Then
         hRet("status") = -1
         hRet("description") = "File doesn't exist"
      Else
         Call hRet.Add ("status", 0)
         Call hRet.Add ("description", "")

         oXmlOldStructure.Load (in_cs_fileLoc)

         If oXmlOldStructure.parseError.errorCode Then
            hRet("status") = -2
            hRet("description") = "Failed to parse old structure"
         Else
            Dim oNodes: Set oNodes = m_oFileNodes.selectNodes ("./file")
            Dim oNodesOld: Set oNodesOld = oXmlOldStructure.selectNodes ("//file")
            Dim oNode, oNodeOld
            Dim fileId, modifDate, size, fileLoc
   
            For Each oNode In oNodes
               fileLoc = oNode.selectSingleNode ("./location").Text

               If isExcludedFile (fileLoc) Or isExcludedPath (fileLoc) Then
                  Call m_oSupport.arrayPush (aExcludedFiles, fileLoc)
               Else
                  fileId = oNode.getAttribute ("ID")

                  Set oNodeOld = oXmlOldStructure.selectSingleNode ("//file[@ID='"& fileId &"']")
                  If oNodeOld is Nothing Then
                     Call m_oSupport.arrayPush (aNewFiles, fileLoc)

                     hRet("status") = hRet("status") + 1
                  Else
                     If oNode.selectSingleNode ("./modifiedDate").Text <> oNodeOld.selectSingleNode ("./modifiedDate").Text _
                        Or oNode.selectSingleNode ("./size").Text <> oNodeOld.selectSingleNode ("./size").Text _
                     Then
                        If oNode.selectSingleNode ("./size").Text = oNodeOld.selectSingleNode ("./size").Text _
                           And Abs (DateDiff ("n", oNode.selectSingleNode ("./modifiedDate").Text, oNodeOld.selectSingleNode ("./modifiedDate").Text)) = 60 _
                        Then
                           ' hack for sync between comp with diff OS -- file hasn't been modified on both comps, but file's modify date is different for exactly ONE HOUR !!!
                        Else
                           Call m_oSupport.arrayPush (aModifFiles, fileLoc)
                                                 
                           hRet("status") = hRet("status") + 1
                        End If
                     End If
                  End If
               End If
            Next

            For Each oNodeOld In oNodesOld
               fileLoc = oNodeOld.selectSingleNode ("./location").Text

               If isExcludedFile (fileLoc) Or isExcludedPath (fileLoc) Then
                  Call m_oSupport.arrayPush (aExcludedFiles, fileLoc)
               Else
                  fileId = oNodeOld.getAttribute ("ID")

                  Set oNode = m_oXmlFileStructure.selectSingleNode ("//file[@ID='"& fileId &"']")
                  If oNode is Nothing Then
                     Call m_oSupport.arrayPush (aDelFiles, fileLoc)

                     hRet("status") = hRet("status") + 1
                  End If
               End If
            Next

            Set oNodes = Nothing
            Set oNodesOld = Nothing
         End If
      End If

      Call hRet.Add ("newFiles", aNewFiles)
      Call hRet.Add ("modifFiles", aModifFiles)
      Call hRet.Add ("delFiles", aDelFiles)
      Call hRet.Add ("exclFiles", aExcludedFiles)

      Set compareStructure = hRet

      Set oXmlOldStructure = Nothing
   End Function


   ' #####
   ' ## Class::CFileStructure::generateFileId (ByVal in_gfi_fileLoc)
   ' ## return: generated tag name from file location
   Public Function generateFileId (ByVal in_gfi_fileLoc) ' return: generated node id from file location
      in_gfi_fileLoc = Replace (in_gfi_fileLoc, "\", "_")
      in_gfi_fileLoc = Replace (in_gfi_fileLoc, "'", "")

      generateFileId = UCase (in_gfi_fileLoc)
   End Function


   ' #####
   ' ## Class::CFileStructure::generateTagName (ByVal in_gtn_fileLoc)
   ' ## return: generated tag name from file location
   Public Function generateTagName (ByVal in_gtn_fileLoc) ' return: generated tag name from file location
      Dim fileId: fileId = in_gtn_fileLoc

      fileId = Replace (fileId, "-", "")
      fileId = Replace (fileId, "+", "")
      fileId = Replace (fileId, "/", "")
      fileId = Replace (fileId, "*", "")
      fileId = Replace (fileId, "=", "")

      fileId = Replace (fileId, "[", "")
      fileId = Replace (fileId, "]", "")
      fileId = Replace (fileId, "(", "")
      fileId = Replace (fileId, ")", "")
      fileId = Replace (fileId, "{", "")
      fileId = Replace (fileId, "}", "")
      fileId = Replace (fileId, " ", "")
      fileId = Replace (fileId, ",", "")
      fileId = Replace (fileId, ";", "")
      fileId = Replace (fileId, "~", "")
      fileId = Replace (fileId, "&", "")
      fileId = Replace (fileId, "!", "")
      fileId = Replace (fileId, "?", "")
      fileId = Replace (fileId, "@", "")
      fileId = Replace (fileId, "#", "")
      fileId = Replace (fileId, "$", "")
      fileId = Replace (fileId, "'", "")
      fileId = Replace (fileId, """", "")

      fileId = Replace (fileId, ":", "")
      fileId = Replace (fileId, "\", "_")
      fileId = Replace (fileId, ".", "_")

      generateTagName = fileId
   End Function


   ' #####
   ' ## Class::CFileStructure::getFileById (ByVal in_gfbi_fileId) ' id - file location
   ' ## return: xml file node
   ' ##
   ' ## !!! Help !!! 
   ' ## to get file node by id ' attrib name & val are case sensitive
   Public Function getFileById (ByVal in_gfbi_fileId) ' return: xml file node
      If m_oXmlFileStructure is Nothing Then
         Err.Description = "setStartingPoint/populateStructure needs to be executed first."
         Err.Raise (19003)
      End If

      Dim fileNodes: Set fileNodes = m_oXmlFileStructure.selectNodes ("//file")

      If fileNodes.length = 0 Then
         getFileById = ""
      Else
         Set getFileById = m_oXmlFileStructure.selectSingleNode ("//file[@ID='"& UCase (in_gfbi_fileId) &"']")
      End If
   End Function


   ' #####
   ' ## Class::CFileStructure::getFullyUsedFolders ()
   Public Function getFullyUsedFolders ()
      getFullyUsedFolders = m_aFullyUsedDirs
   End Function

   
   ' #####
   ' ## Class::CFileStructure::getFullyUsedFiles ()
   Public Function getFullyUsedFiles ()
      getFullyUsedFiles = m_aFullyUsedFiles
   End Function


   ' #####
   ' ## Class::CFileStructure::saveStructure (ByVal in_ss_path)
   Public Sub saveStructure (ByVal in_ss_path)
      If m_oXmlFileStructure is Nothing Then
         Err.Description = "setStartingPoint/populateStructure needs to be executed first."
         Err.Raise (19003)
      End If

'      Dim oXsl: Set oXsl = WScript.CreateObject ("Microsoft.XMLDOM")
'      oXsl.async = False
'      oXsl.loadXML ("<xsl:stylesheet version='1.0' xmlns:xsl='http://www.w3.org/1999/XSL/Transform'>" &_
'         "<xsl:output method='xml' indent='yes'/>" &_
'         "<xsl:template match='@* | node()'>" &_
'         "<xsl:copy>" &_
'         "<xsl:apply-templates select='@* | node()'/>" &_
'         "</xsl:copy>" &_
'         "</xsl:template>" &_
'         "</xsl:stylesheet>")
'MsgBox oXsl.xml  
'      ' Transform
'
'      Dim oResult: Set oResult = WScript.CreateObject("Microsoft.XMLDOM")
'      Call m_oXmlFileStructure.transformNodeToObject (oXsl, oResult)
'      
'MsgBox oResult.xml
'   Call oResult.transformNodeToObject (oXsl, m_oXmlFileStructure)
'MsgBox m_oXmlFileStructure.xml
      m_oXmlFileStructure.save (m_oShell.ExpandEnvironmentStrings (in_ss_path))
'MsgBox "Here"
   End Sub
End Class

'Option Explicit

'Dim obj: Set obj = new CStringManip
'ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile ("TCPconsulting_debug.vbs").ReadAll

'MsgBox obj.ClassName
'MsgBox obj.cleanPathString ("C:\\dima\hi\\\\buy\")
'MsgBox obj.escapeString ("This[that\")
'MsgBox obj.replaceRegExp ("ThisIs__a game", "is.*g", " is g")
'MsgBox obj.IsThereNonUsAscii ("%UserProfile%\Temp\te st.txt")


Class CStringManip
   Private m_className__

   Private Sub Class_Initialize
      m_className__ = "CStringManip"
   End Sub

   Public Property Get ClassName
      ClassName = m_className__
   End Property


   ' #####
   ' ## Class::CStringManip::cleanPathString (ByVal in_cps_str)
   ' ## return: clean file\directory path
   Public Function cleanPathString (ByVal in_cps_str) ' return: clean file\directory path
      Dim fn: fn = "cleanPathString"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

      in_cps_str = Trim (in_cps_str)

      in_cps_str = replaceRegExp (in_cps_str, "\\+", "\")
      cleanPathString = replaceRegExp (in_cps_str, "\\$", "")

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Function


   ' #####
   ' ## Class::CStringManip::replaceRegExp (ByVal in_rre_origStr, ByVal in_rre_pattern, ByVal in_rre_replStr)
   Public Function replaceRegExp (ByVal in_rre_origStr, ByVal in_rre_pattern, ByVal in_rre_replStr)
      Dim oRegExp, str

      Dim fn: fn = "replaceRegExp"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

      Set oRegExp = New RegExp

      With oRegExp
         .Pattern = in_rre_pattern
         .IgnoreCase = True
         .Global = True
      End With

      replaceRegExp = oRegExp.Replace (in_rre_origStr, in_rre_replStr)

      Set oRegExp = Nothing

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Function


   ' #####
   ' ## Class::CStringManip::escapeString (ByVal in_es_str)
   ' ## return: escaped string
   Public Function escapeString (ByVal in_es_str) ' return: escaped string
      Dim fn: fn = "escapeString"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

      in_es_str = Replace (in_es_str, "\", "\\")
      in_es_str = Replace (in_es_str, "(", "\(")
      in_es_str = Replace (in_es_str, ")", "\)")
      in_es_str = Replace (in_es_str, "[", "\[")
      in_es_str = Replace (in_es_str, "]", "\]")
      in_es_str = Replace (in_es_str, "{", "\{")
      in_es_str = Replace (in_es_str, "}", "\}")
      in_es_str = Replace (in_es_str, ".", "\.")
      in_es_str = Replace (in_es_str, "+", "\+")

      escapeString = in_es_str

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Function


   ' #####
   ' ## Class::CStringManip::lpad (ByVal in_lp_str, ByVal in_lp_newLen, ByVal in_lp_chPad)
   ' ## return: left padded string
   Public Function lpad (ByVal in_lp_str, ByVal in_lp_newLen, ByVal in_lp_chPad) ' return: left padded string
      Dim fn: fn = "lpad"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

      lpad = Right (String (in_lp_newLen, in_lp_chPad) & in_lp_str, in_lp_newLen)

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Function

   
   ' #####
   ' ## Class::CStringManip::rpad (ByVal in_rp_str, ByVal in_rp_newLen, ByVal in_rp_chPad)
   ' ## return: right padded string
   Public Function rpad (ByVal in_rp_str, ByVal in_rp_newLen, ByVal in_rp_chPad) ' return: right padded string
      Dim fn: fn = "rpad"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

      rpad = Left (in_rp_str & String (in_rp_newLen, in_rp_chPad), in_rp_newLen)

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Function


   ' #####
   ' ## Class::CStringManip::getBoolValue (ByVal in_gbv_val)
   ' ## return: converted boolean value
   Public Function getBoolValue (ByVal in_gbv_val) ' return: converted boolean value
      Dim fn: fn = "getBoolValue"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

      getBoolValue = False

      If in_gbv_val = True Then
         getBoolValue = True
      ElseIf StrComp (in_gbv_val, "true", vbTextCompare) = 0 Then
         getBoolValue = True
      End If

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Function


   ' #####
   ' ## Class::CStringManip::Bin2Dec (ByVal in_b2d_val)
   ' ## return: decimal for binary value
   Public Function Bin2Dec (ByVal in_b2d_val) ' return: decimal for binary value
      Dim fn: fn = "Bin2Dec"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

      Dim bitsCnt: bitsCnt = Len (in_b2d_val)
      Dim bitIndx, tmp

      For bitIndx = 1 To bitsCnt
         tmp = tmp + Mid (in_b2d_val, bitsCnt - bitIndx + 1, 1) * (2 ^ (bitIndx - 1))
      Next

      Bin2Dec = tmp

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Function


   ' #####
   ' ## Class::CStringManip::Dec2Bin (ByVal in_d2b_val)
   ' ## return: binary for Dec value
   Public Function Dec2Bin (ByVal in_d2b_val) ' return: binary for Dec value
      Dim tmp

      Dim fn: fn = "Dec2Bin"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

      tmp = Trim (in_d2b_val Mod 2)
      in_d2b_val = in_d2b_val \ 2

      Do While in_d2b_val <> 0
         tmp = Trim (in_d2b_val Mod 2) & tmp
         in_d2b_val = in_d2b_val \ 2
      Loop

      Dec2Bin = tmp

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Function


   ' #####
   ' ## Class::CStringManip::Hex2Dec (ByVal in_h2d_val)
   ' ## return: decimal for Hex value
   ' !!! HELP !!!
   ' ## Any to Hex - Hex (val)
   Public Function Hex2Dec (ByVal in_h2d_val) ' return: decimal for Hex value
      Dim fn: fn = "Hex2Dec"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

      in_h2d_val = Trim (in_h2d_val)

      If InStr (1, in_h2d_val, "&H", vbTextCompare) <> 1 Then in_h2d_val = "&H"& in_h2d_val ' &O - Oct

      Hex2Dec = CLng (in_h2d_val)

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Function


   ' #####
   ' ## Class::CStringManip::Hex2Bin (ByVal in_h2b_val)
   ' ## return: binary for Hex value
   Public Function Hex2Bin (ByVal in_h2b_val) ' return: binary for Hex value
      Dim fn: fn = "Hex2Bin"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

      Hex2Bin = Dec2Bin (Hex2Dec (in_h2b_val))
   
      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Function


   ' #####
   ' ## Class::CStringManip::IsThereNonUsAscii (ByVal in_itna_str)
   ' ## return: True || False
   Public Function IsThereNonUsAscii (ByVal in_itna_str) ' return: True || False
      Dim oRegExp, colMatches

      Dim fn: fn = "IsThereNonUsAscii"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

      IsThereNonUsAscii = False
      
      Set oRegExp = New RegExp
      oRegExp.Pattern = "[^A-Za-z 0-9 \.,\?'""!@#\$%\^&\*\(\)-_=\+;\}\{\[\]`~]"
      oRegExp.Global = True
      
      Set colMatches = oRegExp.Execute (in_itna_str)

      If colMatches.Count > 0 Then IsThereNonUsAscii = True
   
      Set oRegExp = Nothing

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Function
End Class

'Option Explicit
'
'Dim obj: Set obj = new CSupport
'ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile ("TCPconsulting_debug.vbs").ReadAll

'MsgBox obj.ClassName
'MsgBox obj.getScriptName ()
'MsgBox obj.getMyTimeStamp ()
'MsgBox obj.getScriptParentProcessInfo ().ProcessId
'Dim users, aKeys
'MsgBox obj.getCurrUserId (users)
'MsgBox "User: "& Join (users.Keys, vbcrlf) & vbcrlf &"Domain: "& Join (users.Items, vbcrlf)
'aKeys = users.Keys
'MsgBox aKeys(0)

Class CSupport
   Private m_className__

   Private Sub Class_Initialize
      m_className__ = "CSupport"
   End Sub

   Public Property Get ClassName
      ClassName = m_className__
   End Property

   ' #####
   ' ## Class::CSupport::getScriptName ()
   ' ## return: script name for .vbs, .vbe files
   Public Function getScriptName () ' return: script name for .vbs, .vbe files
      Dim scriptName
      
      scriptName = WScript.ScriptName

      getScriptName = Replace (scriptName, ".vbs", "")
      getScriptName = Replace (getScriptName, ".vbe", "")
   End Function


   ' #####
   ' ## Class::CSupport::getMyTimeStamp ()
   ' ## return: MyTimeStamp [YYYMMDDhhmm]
   Public Function getMyTimeStamp () ' return: MyTimeStamp [YYYMMDDhhmm]
      Dim fn: fn = "getMyTimeStamp"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

      Dim currDate: currDate = Now
      Dim currMonth: currMonth = DatePart ("m" , currDate)
      Dim currDay: currDay = DatePart ("d" , currDate)

      If currMonth < 10 Then currMonth = "0"& currMonth

      If currDay < 10 Then currDay = "0"& currDay
      
      getMyTimeStamp = DatePart ("yyyy" , currDate) & currMonth & currDay & Replace (FormatDateTime (Time,4),":","")

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Function


   ' #####
   ' ## Class::CSupport::getScriptParentProcessInfo ()
   ' ## return: Parent Process object || Nothing
   '
   ' NIU Function muGetScriptParentProcessInfo (vbsScriptLoc)
   ' NIU    Dim fn: fn = "muGetScriptParentProcessInfo"
   ' NIU    If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)
   ' NIU 
   ' NIU    Set oWmi = GetObject ("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
   ' NIU    Set colWscriptProcs = oWmi.ExecQuery ("SELECT * FROM Win32_Process Where Name='wscript.exe'")
   ' NIU 
   ' NIU    For Each oWscriptProc In colWscriptProcs
   ' NIU       If InStr (oWscriptProc.CommandLine, vbsScriptLoc) > O Then 
   ' NIU          Set colProcs = oWmi.ExecQuery  ("SELECT * FROM Win32_Process Where ProcessId='"& oWscriptProc.ParentProcessId &"'")
   ' NIU 
   ' NIU          For Each muGetScriptParentProcessInfo in colProcs
   ' NIU             If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   ' NIU             Exit Function
   ' NIU          Next
   ' NIU       End If
   ' NIU    Next
   ' NIU 
   ' NIU    If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   ' NIU End Function
   Public Function getScriptParentProcessInfo () ' return: Parent Process object || Nothing
      Dim fn: fn = "getScriptParentProcessInfo"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

      Dim oWmi: Set oWmi = GetObject ("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
      Dim colWscriptProcs: Set colWscriptProcs = oWmi.ExecQuery ("SELECT * FROM Win32_Process Where Name='wscript.exe'")
      Dim oWscriptProc

      For Each oWscriptProc In colWscriptProcs
         If InStr (oWscriptProc.CreationDate, getMyTimeStamp ()) = 1 Then 
            Set getScriptParentProcessInfo = GetObject ("winmgmts:root\cimv2:Win32_Process.Handle='"& oWscriptProc.ParentProcessId &"'")
         End If
      Next

      Set oWmi = Nothing
      Set colWscriptProcs = Nothing
      Set oWscriptProc = Nothing

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Function


   ' #####
   ' ## Class::CSupport::getCurrUserId (ByRef in_ref_gcui_colUsers)
   ' return: 0 - using Env variable || 1 - using WScript.Network object || 2 - using running explorer.exe instances || -1 - failed
   Public Function getCurrUserId (ByRef in_ref_gcui_colUsers) ' return: 0 - using Env variable || 1 - using WScript.Network object || 2 - using running explorer.exe instances || -1 - failed
      Dim userId, domain

      Dim fn: fn = "getCurrUserId"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

      getCurrUserId = -1

      Set in_ref_gcui_colUsers = CreateObject ("Scripting.Dictionary")
      Dim oShell: Set oShell = WScript.CreateObject ("WScript.Shell")

      Dim procEnv: Set procEnv = oShell.Environment ("Process")
      userId = procEnv.Item ("USERNAME")
      getCurrUserId = 0

      If userId = "" Then
         getCurrUserId = 1

         Dim oNet: Set oNet = WScript.CreateObject ("WScript.Network")
         userId = oNet.UserName
         domain = oNet.UserDomain

         Set oNet = Nothing
      End If

      If userId = "" Then
         getCurrUserId = 2

         Dim comp: comp = "."
         Dim oWmi: Set oWmi = GetObject ("winmgmts:{impersonationLevel=impersonate}!\\"& comp &"\root\cimv2")
         Dim colProcs: Set colProcs = oWmi.ExecQuery ("SELECT Name FROM Win32_Process WHERE Name='explorer.exe'")
         Dim oProc, tmp

         For Each oProc In colProcs
            tmp = oProc.GetOwner (userId, domain)

            If tmp = 0 And Not in_ref_gcui_colUsers.Exists (userId) Then 
               in_ref_gcui_colUsers.Add userId, domain
            End If
         Next
      Else
         in_ref_gcui_colUsers.Add userId, domain
      End If

      If in_ref_gcui_colUsers.Count = 0 Then getCurrUserId = -1
      
      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Function


   ' #####
   ' ## Class::CSupport::getCompName ()
   ' ## return: ComputerName
   Public Function getCompName () ' return: ComputerName
      Dim oNet

      Dim fn: fn = "getCompName"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)
      getCompName = ""

      Set oNet = CreateObject ("WScript.Network")
      getCompName = oNet.ComputerName
      Set oNet = Nothing

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Function


   ' #####
   ' ## Class::CSupport::touchFile (ByVal in_tf_path, ByVal in_tf_file, ByVal in_tf_date)
   Public Sub touchFile (ByVal in_tf_path, ByVal in_tf_file, ByVal in_tf_date)
     
      Dim fn: fn = "touchFile"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

      Dim oApp: Set oApp = CreateObject("Shell.Application") 
      Dim oFolder: Set oFolder = oApp.NameSpace (in_tf_path) 
      Dim oFile: Set oFile = oFolder.ParseName (in_tf_file) 

      oFile.ModifyDate = in_tf_date 

      Set oFile = Nothing 
      Set oFolder = Nothing 
      Set oApp = Nothing 

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Sub


   ' #####
   ' ## Class::CSupport::readFileToArray (ByVal in_rfta_oFso, ByVal in_rfta_fileLoc, ByRef in_ref_rfta_aLines, ByRef in_ref_rfta_rowCnt)
   ' ## return: "" - ok || Error Msg;
   Public Function readFileToArray (ByVal in_rfta_oFso, ByVal in_rfta_fileLoc, ByRef in_ref_rfta_aLines, ByRef in_ref_rfta_rowCnt) ' return: "" - ok || Error Msg;
      Dim indx: indx = 0
      readFileToArray = ""

      On Error Resume Next

      Dim fn: fn = "readFileToArray"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

      Dim oStream: Set oStream = in_rfta_oFso.OpenTextFile (in_rfta_fileLoc, 1)

      If Err Then
      Else
         Do Until oStream.AtEndOfStream
            Redim Preserve in_ref_rfta_aLines(indx)

            in_ref_rfta_aLines(indx) = oStream.ReadLine
            indx = indx + 1
         Loop

        oStream.Close
      End If

      in_ref_rfta_rowCnt = indx

      If Err Then
         readFileToArray = Err.Number &": "& Err.Description &" ["& in_rfta_fileLoc  &"]"
      End If

      Set oStream = Nothing
      On Error GoTo 0

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Function


   ' #####
   ' ## Class::CSupport::arrayPush (ByRef in_ap_arr, ByVal in_ap_item)
   Public Sub arrayPush (ByRef in_ap_arr, ByVal in_ap_item)
      Dim aItems, item

      If IsArray (in_ap_item) Then
         aItems = in_ap_item
      Else
         aItems = Array (in_ap_item)
      End If

      For Each item In aItems
         Redim Preserve in_ap_arr (UBound (in_ap_arr) + 1)
         in_ap_arr (UBound (in_ap_arr)) = item
      Next
   End Sub


   ' #####
   ' ## Class::CSupport::makeDir (ByVal in_md_oFso, ByVal in_md_path)
   Public Sub makeDir (ByVal in_md_oFso, ByVal in_md_path)
      Dim parentPath

      Dim fn: fn = "Fatal error: cannot create path """& in_md_path &""" @ makeDir"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

      If Not in_md_oFso.FolderExists (in_md_path) Then
         parentPath = in_md_oFso.GetParentFolderName (in_md_path)

         If Not in_md_oFso.FolderExists (parentPath) And parentPath <> "" Then Call makeDir (in_md_oFso, parentPath)
         If Not in_md_oFso.FolderExists (in_md_path) Then in_md_oFso.CreateFolder (in_md_path)
      
      End If

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Sub


   ' #####
   ' ## Class::CSupport::sendMail (ByVal in_sm_emailFrom, ByVal in_sm_emailTo, ByVal in_sm_subject, ByVal in_sm_bodyTxt, ByVal in_sm_bodyHtml, ByVal in_sm_attach, ByVal in_sm_smtpServer, ByVal in_sm_smtpUser, ByVal in_sm_smtpPass, ByVal in_sm_smtpPort)
   ' ## return: "" || error description
   Public Function sendMail (ByVal in_sm_emailFrom, ByVal in_sm_emailTo, ByVal in_sm_subject, ByVal in_sm_bodyTxt, ByVal in_sm_bodyHtml, ByVal in_sm_attach, ByVal in_sm_smtpServer, ByVal in_sm_smtpUser, ByVal in_sm_smtpPass, ByVal in_sm_smtpPort) ' return: "" || error description
      Dim indx, oEmail

      Dim fn: fn = "sendMail"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

      sendMail = ""
      
      On Error Resume Next

      Set oEmail = CreateObject( "CDO.Message" ) ' using CDOSYS

      With oEmail
         .From     = in_sm_emailFrom
         .To       = in_sm_emailTo
         ' .Cc     = 
         ' .Bcc    = 
         .Subject  = in_sm_subject
         
         If in_sm_bodyHtml = "" Then .TextBody = in_sm_bodyTxt Else .HTMLBody = in_sm_bodyHtml

         'oEmail.CreateMHTMLBody "url"
         'oEmail.CreateMHTMLBody "file://c|/temp/body.htm"

         If IsArray( in_sm_attach ) Then
            For indx = 0 To UBound( in_sm_attach )
               .AddAttachment Replace (in_sm_attach(indx), "\", "\\" ),"",""
            Next
         ElseIf in_sm_attach <> "" Then
            .AddAttachment Replace (in_sm_attach, "\", "\\" ),"",""
         End If

         If in_sm_smtpPort = "" Then in_sm_smtpPort = 25

         With .Configuration.Fields
            .Item( "http://schemas.microsoft.com/cdo/configuration/sendusing"      ) = 2
            .Item( "http://schemas.microsoft.com/cdo/configuration/smtpserver"     ) = in_sm_smtpServer
            .Item( "http://schemas.microsoft.com/cdo/configuration/smtpserverport" ) = in_sm_smtpPort
            .Item( "http://schemas.microsoft.com/cdo/configuration/sendusername"   ) = in_sm_smtpUser
            .Item( "http://schemas.microsoft.com/cdo/configuration/sendpassword"   ) = in_sm_smtpPass

            'http://schemas.microsoft.com/cdo/configuration/smtpauthenticate
            'Const cdoAnonymous = 0 'Do not authenticate
            'Const cdoBasic = 1 'basic (clear-text) authentication
            'Const cdoNTLM = 2 'NTLM

            'http://schemas.microsoft.com/cdo/configuration/smtpusessl = True
            'http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout = 60

            .Update
         End With

        .Send
      End With

      If Err Then
         sendMail = "ERROR " & Err.Number &": "& Err.Description
         Err.Clear
      End If

      Set oEmail = Nothing

      On Error Goto 0

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Function


   ' #####
   ' ## Class::CSupport::countRunningAppInstances (ByVal in_crai_appName)
   ' ## return: count
   Public Function countRunningAppInstances (ByVal in_crai_appName) ' return: count
      Dim fn: fn = "countRunningAppInstances"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

      countRunningAppInstances = 0

      If in_crai_appName = "" Then Exit Function

      On Error Resume Next

      Dim oWmiService: Set oWmiService = GetObject ("winmgmts:")
      Dim cnt: cnt = 0
      Dim item

      in_crai_appName = "\"& UCase (in_crai_appName)
      For Each item In oWmiService.InstancesOf ("Win32_process")
         If InStr (UCase (item.CommandLine), in_crai_appName) Then
            If Err Then Exit Function ' for OS where .CommandLine is not supported
            cnt = cnt + 1
         ElseIf item.CommandLine = "" And InStr (UCase (item.ExecutablePath), in_crai_appName) Then
            cnt = cnt + 1
         End If
      Next

      Set oWmiService = Nothing

      On Error Goto 0

      countRunningAppInstances = cnt

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)
   End Function


   ' #####
   ' ## Class::CSupport::getFolderNamespace (ByVal in_gf_path)
   ' ## return: folder object
   Public Function getFolderNamespace (ByVal in_gf_path) ' return: folder object
      Dim oApp: Set oApp = CreateObject ("Shell.Application")
      Dim oFolder, oShell

      On Error Resume Next

      Set getFolderNamespace = oApp.Namespace (in_gf_path)

      Set oFolder = getFolderNamespace.Self
      If Err Then
         Err.Clear

         Set oShell = CreateObject ("WScript.Shell")
         in_gf_path = oShell.ExpandEnvironmentStrings (in_gf_path)
         Set getFolderNamespace = oApp.Namespace (in_gf_path)
      End If

      On Error Goto 0

      Set oApp = Nothing
      Set oShell = Nothing
      Set oFolder = Nothing
   End Function


   ' #####
   ' ## Class::CSupport::deleteCookies ()
   ' ## return: path to Internet Cookies folder
   Public Function deleteCookies (ByVal in_dc_folderPath, ByVal in_dc_bDelSubFolder, ByVal in_dc_bDelHidden, ByRef in_ref_dc_cntFileFail, ByRef in_ref_dc_cntFolderFail) ' return: path to Internet Cookies folder
      Dim fn: fn = "deleteCookies"
      If gl_tcpc_debug = True Then tcpc_debugStackAdd (fn)

'constants for special folders in windows: 
'&H1& Internet Explorer
'&H2& Programs
'&H3& Control Panel
'&H4& Printers and Faxes
'&H5& My Documents *
'&H6& Favorites *
'&H7& Startup *
'&H8& My Recent Documents *
'&H9& SendTo
'&Ha& Recycle Bin *
'&Hb& Start Menu
'&Hd& My Music *
'&He& My Videos *
'&H10& Desktop
'&H11& My Computer
'&H12& My Network Places
'&H13& NetHood
'&H14& Fonts
'&H15& Templates
'&H16& All Users Start Menu
'&H17& All Users Programs
'&H18& All Users Startup
'&H19& All Users Desktop
'&H1a& Application Data
'&H1b& PrintHood
'&H1c& Local Settings\Application Data *
'&H19& All Users Favorites
'&H20& Local Settings\ Temporary Internet Files *
'&H21& Cookies *
'&H22& Local Settings\History *
'&H23& All Users Application Data
'&H24& Windows
'&H25& System32
'&H26& Program Files
'&H27& My Pictures *
'&H28& User Profile
'&H2b& Common Files
'&H2e& All Users Templates
'&H2f& Administrative Tools
'&H31& Network Connections

      Const cFolderCookies = &H21&

      If (in_dc_folderPath = "") Then in_dc_folderPath = cFolderCookies

      deleteCookies = in_dc_folderPath

      On Error Resume Next

      Dim oFolderNS: Set oFolderNS = getFolderNamespace (in_dc_folderPath)
      Dim oFolder: Set oFolder = oFolderNS.Self

      in_ref_dc_cntFileFail = -1
      in_ref_dc_cntFolderFail = -1

      If Err Then
         Err.Clear
         Exit Function
      End If

      deleteCookies = oFolder.Path

      Dim colItems: Set colItems = oFolderNS.Items
      Dim oItem

      If colItems.Count < 1 And in_dc_bDelHidden <> True Then Exit Function

      Dim oFso: Set oFso = WScript.CreateObject ("Scripting.FileSystemObject")

      in_ref_dc_cntFileFail = 0
      in_ref_dc_cntFolderFail = 0

      If in_dc_bDelHidden = True Then
         Set oFolder = oFso.GetFolder (deleteCookies)

         Set colItems = oFolder.Files
         For Each oItem in colItems
            Call oFso.DeleteFile (oFso.BuildPath (deleteCookies, oItem.Name), True)

            If Err Then
               in_ref_dc_cntFileFail = in_ref_dc_cntFileFail + 1
               Err.Clear
            End If
         Next

         If in_dc_bDelSubFolder = True Then
            Set colItems = oFolder.SubFolders
            For Each oItem in colItems
               Call oFso.DeleteFolder (oFso.BuildPath (deleteCookies, oItem.Name), True)

               If Err Then
                  in_ref_dc_cntFolderFail = in_ref_dc_cntFolderFail + 1
                  Err.Clear
               End If
            Next
         End If
      Else
         For Each oItem in colItems
            If oItem.isFolder Then 
               If in_dc_bDelSubFolder = True Then Call oFso.DeleteFolder (oFso.BuildPath (deleteCookies, oItem.Name), True)
            Else
               Call oFso.DeleteFile (oFso.BuildPath (deleteCookies, oItem.Name), True)
            End If

            If Err Then
               If oItem.isFolder Then
                  in_ref_dc_cntFolderFail = in_ref_dc_cntFolderFail + 1
               Else
                  in_ref_dc_cntFileFail = in_ref_dc_cntFileFail + 1
               End If

               Err.Clear
            End If
         Next
      End If

      On Error Goto 0

      If gl_tcpc_debug = True Then tcpc_debugStackRemove (fn)

      Set oFso = Nothing
      Set oFolderNS = Nothing
      Set oFolder = Nothing
      Set colItems = Nothing
      Set oItem = Nothing
   End Function
End Class

Dim gl_tcpc_debug
Dim gl_tcpc_debugSeparator
Dim gl_tcpc_debugStack: gl_tcpc_debugStack = ""

        
' #####
' ## tcpc_debugStackAdd (ByVal in_dsa_funcName)
Sub tcpc_debugStackAdd (ByVal in_dsa_funcName)
   gl_tcpc_debugStack = gl_tcpc_debugStack & tcpc_getDebugSeparator() & in_dsa_funcName
End Sub


' #####
' ## tcpc_debugStackRemove (ByVal in_dsr_funcName)
Sub tcpc_debugStackRemove (ByVal in_dsr_funcName)
   Dim startPos: startPos = InStrRev (gl_tcpc_debugStack, tcpc_getDebugSeparator() & in_dsr_funcName)

   If startPos > 0 And Len (gl_tcpc_debugStack) = Len (tcpc_getDebugSeparator() & in_dsr_funcName) + startPos - 1 Then gl_tcpc_debugStack = Left (gl_tcpc_debugStack, startPos - 1)
End Sub


' #####
' ## tcpc_debugGetStack ()
Function tcpc_debugGetStack ()
   tcpc_debugGetStack = gl_tcpc_debugStack
End Function


' #####
' ## tcpc_getDebugSeparator ()
Function tcpc_getDebugSeparator ()
   If gl_tcpc_debugSeparator = "" Then
      tcpc_getDebugSeparator = "::"
   Else
      tcpc_getDebugSeparator = gl_tcpc_debugSeparator
   End If
End Function



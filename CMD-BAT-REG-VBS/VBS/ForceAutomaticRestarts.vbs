'--------------------------------------------------------------------------------- 
'The sample scripts are not supported under any Microsoft standard support 
'program or service. The sample scripts are provided AS IS without warranty  
'of any kind. Microsoft further disclaims all implied warranties including,  
'without limitation, any implied warranties of merchantability or of fitness for 
'a particular purpose. The entire risk arising out of the use or performance of  
'the sample scripts and documentation remains with you. In no event shall 
'Microsoft, its authors, or anyone else involved in the creation, production, or 
'delivery of the scripts be liable for any damages whatsoever (including, 
'without limitation, damages for loss of business profits, business interruption, 
'loss of business information, or other pecuniary loss) arising out of the use 
'of or inability to use the sample scripts or documentation, even if Microsoft 
'has been advised of the possibility of such damages 
'--------------------------------------------------------------------------------- 
Option Explicit
On Error Resume Next

Dim objShell
Dim objUpdateSession
Dim objUpdateSearcher
Dim objUpdateResults
Dim UpdateResults

Wscript.Echo "The script is running..."
Set objShell = CreateObject("Wscript.Shell")
Set objUpdateSession = CreateObject("Microsoft.Update.Session")
Set objUpdateSearcher = objUpdateSession.CreateUpdateSearcher
Set objUpdateResults = objUpdateSearcher.Search("Type='Software'")
Set UpdateResults = objUpdateResults.Updates

Dim UpdateInstalled
Dim UpdateStr
Dim regKeyPath
Dim AutoRebootValue
Dim i

UpdateInstalled = False

For i = 0 to UpdateResults.Count - 1
  UpdateStr = UpdateResults.Item(i).Title
  
  If InStr(UpdateStr,"KB2822241") > 0  Then
    If UpdateResults.Item(i).IsInstalled <> 0 Then
    	UpdateInstalled = True
    End If
  End If
Next

If UpdateInstalled Then
  regKeyPath = "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AlwaysAutoRebootAtScheduledTime"
  AutoRebootValue = objShell.RegRead(regKeyPath)
  
  If Err.Number = 0 Then
    If AutoRebootValue = 1 Then
      WScript.Echo "You have been enabled automatic Windows Update restarts."
    Else
      objShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AlwaysAutoRebootAtScheduledTime","1","REG_DWORD"
    End If
  Else  
    objShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\",""
    objShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\",""
    objShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\AlwaysAutoRebootAtScheduledTime","1","REG_DWORD"
    WScript.Echo "Successfully enabled automatic Windows Update restarts."
  End If
Else
    WScript.Echo "You do not install Windows Update 2822241. To force automatic restarts only after you install Windows Update 2822241."
End If  
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

SetOfflineMode

Sub SetOfflineMode()
    Dim appOutlook
    Dim blnIsOffline
    
    On Error Resume Next
    
    Set appOutlook = CreateObject("Outlook.Application")
    
    ' /* If no error occurred. */
    If Err.Number = 0 Then
        blnIsOffline = appOutlook.Session.Offline
        
        ' /* Enables "Work Offline" only if the current session is online. */
        If Not blnIsOffline Then
            appOutlook.ActiveExplorer.CommandBars.FindControl(, 5613).Execute
            MsgBox "Work Offline has been enabled successfully!", 64, "Tips"
        Else
            MsgBox "Work Offline has been set up!", 48, "Tips"
        End If
    Else
        MsgBox Err.Description, 16, "Error"
    End If
    'appOutlook.Quit
    If Not (appOutlook Is Nothing) Then Set appOutlook = Nothing
End Sub
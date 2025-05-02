'---------------------------------------------------------------------------------
' The sample scripts are not supported under any Microsoft standard support
' program or service. The sample scripts are provided AS IS without warranty
' of any kind. Microsoft further disclaims all implied warranties including,
' without limitation, any implied warranties of merchantability or of fitness for
' a particular purpose. The entire risk arising out of the use or performance of
' the sample scripts and documentation remains with you. In no event shall
' Microsoft, its authors, or anyone else involved in the creation, production, or
' delivery of the scripts be liable for any damages whatsoever (including,
' without limitation, damages for loss of business profits, business interruption,
' loss of business information, or other pecuniary loss) arising out of the use
' of or inability to use the sample scripts or documentation, even if Microsoft
' has been advised of the possibility of such damages.
'---------------------------------------------------------------------------------

Dim pptApp
Dim fd
Dim vntSelectedItem

Set pptApp = CreateObject("PowerPoint.Application")

' Create a FileDialog object as a File Open dialog box.
Set fd = pptApp.FileDialog(3)

'/* To reference the FileDialog object .*/
With fd
    .AllowMultiSelect = True
    .Filters.Clear
    .Filters.Add "All PowerPoint Presentations", "*.pptx,*.ppt,*.pptm,*.ppsx,*.pps,*.ppsm,*.potx,*.pot,*.potm,*.odp"
    .Title = "Select Presentations to Operate"
    
    '/* The user pressed the button .*/
    If .Show = -1 Then
        Dim myPPT
        Dim strResult
        Dim iNow
        Dim iAll
        Dim mySlide
        Dim mySlides
        
        iNow = 0
        iAll = .SelectedItems.Count
        
        '/* Step thru each string in the FileDialogSelectedItems collection. */
        For Each vntSelectedItem In .SelectedItems
            
            ' To reference the opening presentation object.
            Set myPPT = pptApp.Presentations.Open(vntSelectedItem, , , 0)
            ' Set the counter.
            iNow = iNow + 1
            ' Display the progress.
            WScript.Echo CStr(iNow) & "/" & CStr(iAll) & Chr(9) & myPPT.Name
            
            With myPPT
                Set mySlides = .Slides
                
                '/* Step thru each slide. */
                For Each mySlide In mySlides
                    If mySlide.NotesPage.Shapes.Count > 0 Then
                        mySlide.NotesPage.Shapes.Range.Delete
                    End If
                Next
                
                .Save
                .Close
            End With
            
        Next
        
        ' Reminding for task completed in command line.
        WScript.Echo Chr(10) & _
                     "*****************" & Chr(10) & _
                     " Task completed! " & Chr(10) & _
                     "*****************" & Chr(10)
        
        '/* Set the object variable to nothing. */
        Set myPPT = Nothing
        Set mySlide = Nothing
        Set mySlides = Nothing
    End If
    
End With

' Set the object variable to nothing.
Set fd = Nothing
' Exit the created instance of PowerPoint application.
pptApp.Quit
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

Call Main

Sub Main()
    Dim excelApp
    Set excelApp = CreateObject("Excel.Application")
    
    With excelApp.FileDialog(1)
        .AllowMultiSelect = False
        .Title = "Select Access Databases to Operate"
        With .Filters
            .Clear
            .Add "Microsoft Access Databases", "*.mdb,*.accdb"
        End With
        
        If .Show = -1 Then
            Dim strAccessPath
            strAccessPath = .SelectedItems(1)
            
            Dim blnGetError

            ' Call the function to process the selected file.
            blnGetError = TransferAccessToExcel(strAccessPath,excelApp)       
    
            If blnGetError Then
                MsgBox "Runtime error, please check to see if the problems listed below:" & vbNewLine & _
                       "1. ActiveX component (ADO) can't create object;" & vbNewLine & _
                       "2. Provider cannot be found. It may not be properly installed.", 16, "Error"
            Else
                MsgBox "Task has been completed.", 64, "Tips"
            End If
        End If
    End With
End Sub

Function TransferAccessToExcel(ByVal DataSource, Byref ObjectExcel)
    Dim adoConn
    Dim rstSchema
    Dim rstData
    Dim strConn
    Dim strQuery
    Dim strTable
    Dim strCommonConn
    Dim intK
    Dim blnHasError
    Dim qryTable
    Dim wsDest
    Dim wbDest
    Dim colTableNames()
    Dim strFilePath
    Dim strFileName
    Dim intNum
    
    intNum = 0
    
    ' Get the file path from the given path.
    strFilePath = GetFilePathFromFullPath(DataSource)
    ' Get the file name from the given path.
    strFileName = GetFileNameFromFullPath(DataSource)
    
    ' Constants for Access database engine. 
    Const A20072010 = "Provider=Microsoft.ACE.OLEDB.12.0;"
    Const A19972003 = "Provider=Microsoft.Jet.OLEDB.4.0;"
    
    blnHasError = False ' The function returns no error.
    strCommonConn = "Persist Security Info=False;Data Source=" & DataSource
    
    ' Set the Access database engine to Access 2007 - 2010.
    strConn = A20072010 & strCommonConn
    
    On Error Resume Next
    
    ' Try to create Microsoft ActiveX Data Objects (ADO).
    Set adoConn = CreateObject("ADODB.Connection")
    Set rstSchema = CreateObject("ADODB.Recordset")
    Set rstData = CreateObject("ADODB.Recordset")
    
    If Err.Number <> 0 Then
        blnHasError = True
        'Go Sub errExit
    Else
        
        adoConn.Open strConn
        
        If Err.Number <> 0 Then
            Err.Clear
            blnHasError = True
            
            ' Change the Access database engine to Access 1997 - 2003
            strConn = A19972003 & strCommonConn
            adoConn.Open strConn
             
            If Err.Number <> 0 Then
                blnHasError = True
                'GoTo errExit
            Else
                blnHasError = False
            End If
        End If
        
        Set rstSchema = adoConn.OpenSchema(20)  ' adSchemaTables = 20
        
        ' Loop to get all the data table names. 
        Do Until rstSchema.EOF
            If rstSchema("TABLE_TYPE") <> "ACCESS TABLE" And rstSchema("TABLE_TYPE") <> "SYSTEM TABLE" Then
                strTable = rstSchema("TABLE_NAME")
                ReDim Preserve colTableNames(intNum)
                colTableNames(intNum) = strTable
                intNum = intNum + 1
            End If
            
            rstSchema.MoveNext
        Loop
        
        If intNum > 0 Then
            Set wbDest = ObjectExcel.Workbooks.add
            
            For intK = 0 To intNum - 1
                If rstData.State = 1 Then rstData.Close
                
                strQuery = "select * from " & colTableNames(intK)
                
                With rstData
                    .CursorLocation = 3 ' adUseClient = 3
                    .LockType = 3       ' adLockOptimistic = 3
                    .CursorType = 3     ' adOpenStatic = 3
                    .Open strQuery, strConn
                    
                    On Error Resume Next
                    
                    Set wsDest = wbDest.Worksheets(colTableNames(intK))

                    ' Test if destination sheet exists.
                    If Err <> 0 Then
                        Err.Clear
                        ' Insert a new worksheet after the last worksheet in the active workbook.
                        Set wsDest = wbDest.Worksheets.Add(, wbDest.Worksheets(wbDest.Worksheets.Count))
                        ' Rename the added sheet's name as current table's name.
                        wsDest.Name = colTableNames(intK)
                    Else
                        ' Empty cells.
                        wsDest.Cells.Delete xlUp
                    End If
                    
                    Set qryTable = wsDest.QueryTables.Add(rstData, wsDest.Cells(1, 1))
                    
                    ' Show field names from the data source as column headings.
                    qryTable.FieldNames = True
                    qryTable.Refresh
                    ' Don't maintain the connection to the specified data source after the refresh.
                    qryTable.MaintainConnection = False
                    ' Delete the QueryTable object.
                    qryTable.Delete
                End With
            Next
            
            wbDest.SaveAs strFilePath & strFileName
            wbDest.Close
            ObjectExcel.Quit
        End If
    End If
    
    TransferAccessToExcel = blnHasError
    
    ' Close the open recordset to the data tables.
    CloseRst rstSchema
    CloseRst rstData
    
    ' Close the open connection to the database.
    adoConn.Close
    
    ' Release memory.
    If Not rstSchema Is Nothing Then Set rstSchema = Nothing
    If Not rstData Is Nothing Then Set rstData = Nothing
    If Not adoConn Is Nothing Then Set adoConn = Nothing
    If Not ObjectExcel Is Nothing Then Set adoConn = Nothing
    If Not wbDest Is Nothing Then Set adoConn = Nothing
    If Not wsDest Is Nothing Then Set adoConn = Nothing
End Function

' #########################################
' Get file path from a specified full path.
' #########################################
Function GetFilePathFromFullPath(FullPath)
	Dim lngPathSeparatorPosition	' Path separator.
	
	GetFilePathFromFullPath = ""
	lngPathSeparatorPosition = InStrRev(FullPath, "\", -1, 1)
	
	If lngPathSeparatorPosition <> 0 Then GetFilePathFromFullPath = Left(FullPath, lngPathSeparatorPosition)
End Function

' #########################################
' Get file name from a specified full path.
' #########################################
Function GetFileNameFromFullPath(FullPath)
	Dim lngPathSeparatorPosition	' Path separator.
	Dim lngDotPosition				' Dot position.
	Dim strFile						' A full file name.
	
	GetFileNameFromFullPath = ""
	lngPathSeparatorPosition = InStrRev(FullPath, "\", -1, 1)
	
	If lngPathSeparatorPosition <> 0 Then
		strFile = Right(FullPath, Len(FullPath) - lngPathSeparatorPosition)
		lngDotPosition = InStrRev(strFile, ".", -1, 1)
		
		If lngDotPosition <> 0 Then GetFileNameFromFullPath = Left(strFile, lngDotPosition - 1)
	End If
End Function
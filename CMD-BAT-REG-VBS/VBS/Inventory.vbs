'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 2012
'
' NAME:  Manoj Ravikumar Nair
'
' AUTHOR: Microsoft Corp. , Microsoft IT
' DATE  : 6/18/2013
'
' COMMENT: 
'
'==========================================================================

on Error Resume Next 
 
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 
objExcel.Workbooks.Add 
 
objExcel.Cells(1, 1).Value = "  Computer Name  " 
objExcel.Cells(1, 1).Font.Colorindex = 2
objExcel.Cells(1, 1).Font.Bold = True 
objExcel.Cells(1, 1).Interior.ColorIndex = 23 
objExcel.Cells(1, 1).Alignment = "Center" 
Set objRange = objExcel.ActiveCell.EntireColumn 
objRange.AutoFit() 
 
objExcel.Cells(1, 2).Value = "  Manufacturer  " 
objExcel.Cells(1, 2).Font.Colorindex = 2
objExcel.Cells(1, 2).Font.Bold = True 
objExcel.Cells(1, 2).Interior.ColorIndex = 23 
Set objRange = objExcel.Range("B1") 
objRange.Activate 
Set objRange = objExcel.ActiveCell.EntireColumn 
objRange.AutoFit() 
 
objExcel.Cells(1, 3).Value = "  Model  " 
objExcel.Cells(1, 3).Font.Colorindex = 2
objExcel.Cells(1, 3).Font.Bold = True 
objExcel.Cells(1, 3).Interior.ColorIndex = 23 
Set objRange = objExcel.Range("C1") 
objRange.Activate 
Set objRange = ObjExcel.ActiveCell.EntireColumn 
objRange.AutoFit() 
 
objExcel.Cells(1, 4).Value = "  RAM (GB) " 
objExcel.Cells(1, 4).Font.Colorindex = 2
objExcel.Cells(1, 4).Font.Bold = True 
objExcel.Cells(1, 4).Interior.ColorIndex = 23 
Set objRange = objExcel.Range("C1") 
objRange.Activate 
Set objRange = ObjExcel.ActiveCell.EntireColumn 
objRange.AutoFit() 

objExcel.Cells(1, 5).Value = " Operating System "
objExcel.Cells(1, 5).Font.Colorindex = 2
objExcel.Cells(1, 5).Font.Bold = True 
objExcel.Cells(1, 5).Interior.ColorIndex = 23 
Set objRange = objExcel.Range("C1") 
objRange.Activate 
Set objRange = ObjExcel.ActiveCell.EntireColumn 
objRange.AutoFit()

objExcel.Cells(1, 6).Value = " Processor "
objExcel.Cells(1, 6).Font.Colorindex = 2
objExcel.Cells(1, 6).Font.Bold = True 
objExcel.Cells(1, 6).Interior.ColorIndex = 23 
Set objRange = objExcel.Range("C1") 
objRange.Activate 
Set objRange = ObjExcel.ActiveCell.EntireColumn 
objRange.AutoFit()

objExcel.Cells(1, 7).Value = " Drive "
objExcel.Cells(1, 7).Font.Colorindex = 2
objExcel.Cells(1, 7).Font.Bold = True 
objExcel.Cells(1, 7).Interior.ColorIndex = 23 
Set objRange = objExcel.Range("C1") 
objRange.Activate 
Set objRange = ObjExcel.ActiveCell.EntireColumn 
objRange.AutoFit()

objExcel.Cells(1, 8).Value = " Drive Size (GB)"
objExcel.Cells(1, 8).Font.Colorindex = 2
objExcel.Cells(1, 8).Font.Bold = True 
objExcel.Cells(1, 8).Interior.ColorIndex = 23 
Set objRange = objExcel.Range("C1") 
objRange.Activate 
Set objRange = ObjExcel.ActiveCell.EntireColumn 
objRange.AutoFit()

objExcel.Cells(1, 9).Value = " Free Space (GB) "
objExcel.Cells(1, 9).Font.Colorindex = 2
objExcel.Cells(1, 9).Font.Bold = True 
objExcel.Cells(1, 9).Interior.ColorIndex = 23 
Set objRange = objExcel.Range("C1") 
objRange.Activate 
Set objRange = ObjExcel.ActiveCell.EntireColumn 
objRange.AutoFit()
 
 ' Reading the Input File (Text File Containing Computer Names)
y = 2 
Set fso1 = CreateObject("Scripting.FileSystemObject") 
Set pcfile = fso1.OpenTextFile("E:\Code Library\VBS\Computers.txt",1) 
do while Not pcfile.AtEndOfStream 
    computerName = pcfile.readline 
Err.Clear 
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    computerName & "\root\cimv2") 
Set colSettings = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem") 
Set colOSSettings = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
Set colProcSettings = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
Set colDiskSettings = objWMIService.ExecQuery("Select * from Win32_LogicalDisk Where DriveType=3")


If Err.Number = 0 Then 
    For Each objComputer In colSettings 
        strManufacturer = objComputer.Manufacturer 
        strModel = objComputer.Model 
        
        'FormatGB = FormatNumber(size / (1024 * 1024 * 1024), 2,0,0,0)
        strRAM = FormatNumber((objComputer.TotalPhysicalMemory / (1024 * 1024 *1024)), 2,0,0,0)
        
        For Each objOS In colOSSettings
        
        	strOS = objOS.Caption
        	
        For Each objProc In colProcSettings
        
        	strProc = objProc.Name
        	
        	
     
        If computerName <> oldName Then 
          objExcel.Cells(y, 1).Value = computerName 
          objExcel.Cells(y, 1).Alignment = "Center" 
          objExcel.Cells(y, 2).Value = strManufacturer 
          objExcel.Cells(y, 2).Alignment = "Center" 
          objExcel.Cells(y, 3).Value = strModel 
          objExcel.Cells(y, 3).Alignment = "Center"  
          objExcel.Cells(y, 4).Value = strRAM 
          objExcel.Cells(y, 4).Alignment = "Center" 
          objExcel.Cells(y, 5).Value = strOS 
          objExcel.Cells(y, 5).Alignment = "Center" 
          objExcel.Cells(y, 6).Value = strProc 
          objExcel.Cells(y, 6).Alignment = "Center" 
          
          For Each objDisk In colDiskSettings
          
          strDiskDeviceID = objDisk.DeviceID
	          objExcel.Cells(y, 7).Value = strDiskDeviceID 
	          objExcel.Cells(y, 7).Alignment = "Center" 
	          
	          'FormatNumber((objComputer.TotalPhysicalMemory / (1024 * 1024 *1024)), 2,0,0,0)
	          
          	strDiskSize = FormatNumber((objDisk.Size / (1024 * 1024 * 1024)),2,0,0,0)
          	objExcel.Cells(y, 8).Value = strDiskSize 
          	objExcel.Cells(y, 8).Alignment = "Center"           
          strDiskFreeSpace = FormatNumber((objDisk.FreeSpace / (1024 * 1024 * 1024)),2,0,0,0)
          	objExcel.Cells(y, 9).Value = strDiskFreeSpace 
          	objExcel.Cells(y, 9).Alignment = "Center" 
          	
          	y = y + 1
          	
          	
          	Next  
          
          
          
          
          
          
          
          oldName = computerName 
          y = y + 1 
        End If 
       computerName = "" 
       strManufacturer = "" 
       strModel = "" 
       strRAM = "" 
       Err.clear 
    Next 
    
    Next
    
    Next
Else 
         objExcel.Cells(y, 1).Value = computerName 
         objExcel.Cells(y, 1).Alignment = "Center" 
         objExcel.Cells(y, 2).Value = "Not on line" 
         objExcel.Cells(y, 2).Alignment = "Center" 
         y = y + 1 
         Err.clear  
End If 
Loop 
y = y + 1 
objExcel.Cells(y, 1).Value = "Scan Complete" 
objExcel.Cells(y, 1).Font.Bold = True 
 
et objRange = objExcel.Range("A1") 
objRange.Activate 
Set objRange = objExcel.ActiveCell.EntireColumn 
objRange.Autofit() 
Set objRange = objExcel.Range("B1") 
objRange.Activate 
Set objRange = objExcel.ActiveCell.EntireColumn 
objRange.Autofit()  
Set objRange = objExcel.Range("C1") 
objRange.Activate 
Set objRange = objExcel.ActiveCell.EntireColumn 
objRange.Autofit() 
Set objRange = objExcel.Range("D1") 
objRange.Activate 
Set objRange = objExcel.ActiveCell.EntireColumn 
objRange.Autofit() 
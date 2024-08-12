dim wshShell, filesys, newfolder
set wshShell = WScript.CreateObject( "WScript.Shell" )
set filesys = CreateObject("Scripting.FileSystemObject")

Dim intYN, strComputer, objWMIService, colInstalledPrinters, objPrinter, objNetwork
dim Servers
dim strLogDir
dim files, folder

strComputer = "."
   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

   Set colInstalledPrinters =  objWMIService.ExecQuery _
       ("Select * from Win32_Printer Where Network = true")

   For Each objPrinter in colInstalledPrinters
      objPrinter.Delete_
   Next
   


WScript.Quit

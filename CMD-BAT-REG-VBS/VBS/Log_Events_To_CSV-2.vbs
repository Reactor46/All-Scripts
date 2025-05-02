'''' ' Another Logn Script to save evenrts to CSV
On Error Resume Next

''' Defining the variables 
Dim wsNetwork 
Dim outUser 
Dim outComputer 
Dim outIP 
Dim septxt 
Dim logPathname

'' Setting the variables 
' septxt is the column separator (i.e. comma for CSV, tab for tab-delimited, etc) 
septxt = ", "

' logPathname is the filepath to save to. \\server\share\file.csv and C:\file.csv both work if writable by the currently logged-on user. 
logPathname = "\\LOGGINGSERVERNAME\LOGFOLDER\LOGFILE.csv"

''''''''''''''''''''''''''''''''''''''' 
' Computer Name Section 
'''''''''''''''''''''''''''''''''''''''

Set objComputer = CreateObject("Shell.LocalMachine") 
outComputer = objComputer.MachineName

''''''''''''''''''''''''''''''''''''''' 
' User Name Section 
'''''''''''''''''''''''''''''''''''''''

set wsNetwork = createobject("WSCRIPT.Network") 
outUser=wsNetwork.UserName

''''''''''''''''''''''''''''''''''''''' 
' IP Address Section 
'''''''''''''''''''''''''''''''''''''''

strComputer = "." 
Set objWMIService = GetObject("winmgmts:" _ 
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set IPConfigSet = objWMIService.ExecQuery _ 
("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE") 

outIP = ""

For Each IPConfig in IPConfigSet 
If Not IsNull(IPConfig.IPAddress) Then 
For i=LBound(IPConfig.IPAddress) to UBound(IPConfig.IPAddress) 
outIP = outIP & IPConfig.IPAddress(i) & septxt 
Next 
End If 
Next

''''''''''''''''''''''''''''''''''''''' 
' Output Section 
'''''''''''''''''''''''''''''''''''''''

Const ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objTextFile = objFSO.OpenTextFile (logPathname, ForAppending, True)

objTextFile.WriteLine(FormatDateTime(Now(), vbGeneralDate) & septxt & outUser & septxt & outComputer & septxt & outIP)

objTextFile.Close

''''' End Logon sccript to CSV
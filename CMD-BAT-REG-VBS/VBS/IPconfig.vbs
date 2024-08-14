On Error Resume Next

Const ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile ("C:\temp\IPconfigServer.csv", ForAppending, true)

strComputer = "." 
Set objWMIService = GetObject("winmgmts:" _ 
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
 
Set IPConfigSet = objWMIService.ExecQuery _ 
    ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE") 
  
For Each IPConfig in IPConfigSet 
    Output=IPConfig.DNSHostname
    If Not IsNull(Ipconfig.DNSServerSearchOrder) then
	For i=lbound(Ipconfig.DNSServerSearchOrder) to ubound(Ipconfig.DNSServerSearchOrder)
		output = output & "," & Ipconfig.DNSServerSearchOrder(i)
		
	Next
    end if
    objTextFile.WriteLine (Output)
    wscript.echo output
 Next 

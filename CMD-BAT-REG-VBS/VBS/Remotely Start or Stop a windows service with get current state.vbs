dim svcName, sStart, sStop 
dim service  
dim objService 
dim SvrName 
dim input 
dim input2 
Dim svclist
Dim compname
dim objFile
Set WshShell = WScript.CreateObject("WScript.Shell")

on error resume next
 
input = Inputbox ("Computer Name that you would like to Start or stop service on:"& vbCRLF & "Default is this pc", _ 
    "Stop or Start service","localhost") 
	
	If input = "" Then
  WScript.Quit
End If

returnvalue=MsgBox ("Do you want display a list of services on remote pc?"& vbCRLF & "List will show the Service Short Name, the current State and the Service Display Name", 36, "Show Services")
if returnvalue = 6 then 


Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    ("c:\services list.txt", ForWriting, True)
Set colServices =  GetObject("winmgmts://" & input).ExecQuery _
    ("Select * from Win32_Service")
	objTextFile.WriteLine("Short Name") & VbTab & ("Current State") & VbTab & VbTab & ("Display Name")
For Each objService in colServices    
    objTextFile.WriteLine(objService.Name  & VbTab  & _
        objService.State & VbTab & VbTab & objService.DisplayName)
Next
objTextFile.Close
WScript.Sleep 500 
wShshell.Run "notepad c:\services list.txt"
WScript.Sleep 500 

If input = "" Then
 WScript.Quit
End If

	
input2 = Inputbox ("Enter the Service name you want to start or stop:" & vbCRLF & "Use the services list.txt file to get service short name if you don't know it." & vbCRLF & "print spooler = spooler" & vbCRLF & "remote registry = remoteregistry" & vbCRLF & "Windows Event log = eventlog"& vbCRLF & "Default is Print Spooler" & vbCRLF & "","Service Name","spooler")
 
	If input2 = "" Then
  WScript.Quit
End If


SvrName = input 
svcName = input2 
 
Set service = GetObject("winmgmts:!\\" & svrName & "\root\cimv2") 
svc = "Win32_Service=" & "'" & svcName & "'" 
state = inputbox _ 
    ("Enter start or stop to Change the Status of Service"  & vbCRLF & "lowercase only" & vbCRLF & "Default is start","start or stop","start") 
 
If state = "" Then
  WScript.Quit
  End If
 
if state = "start" then 
    Set objService = Service.Get(svc) 
    retVal = objService.StartService() 
	
returnvalue=MsgBox ("Do you want re-open the services list.txt file and verify the state change?", 36, "Started the " & svcName & " service on " & SvrName)
if returnvalue = 6 then 

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    ("c:\services list.txt", ForWriting, True)
Set colServices =  GetObject("winmgmts://" & input).ExecQuery _
    ("Select * from Win32_Service")
	objTextFile.WriteLine("Short Name") & VbTab & ("Current State") & VbTab & ("Display Name")
For Each objService in colServices    
    objTextFile.WriteLine(objService.Name  & VbTab  & _
        objService.State & VbTab  & objService.DisplayName)
Next
objTextFile.Close
WScript.Sleep 500 
wShshell.Run "notepad c:\services list.txt"
WScript.Sleep 500 
end if
WScript.Quit


elseif state = "stop" then 
    Set objService = Service.Get(svc) 
    retVal = objService.StopService() 
end if 
	returnvalue=MsgBox ("Do you want re-open the services list.txt file and verify the state change?", 36, "Stopped the " & svcName & " service on " & SvrName)
if returnvalue = 6 then 

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    ("c:\services list.txt", ForWriting, True)
Set colServices =  GetObject("winmgmts://" & input).ExecQuery _
    ("Select * from Win32_Service")
	objTextFile.WriteLine("Short Name") & VbTab & ("Current State") & VbTab & ("Display Name")
For Each objService in colServices    
    objTextFile.WriteLine(objService.Name  & VbTab  & _
        objService.State & VbTab  & objService.DisplayName)
Next
objTextFile.Close
WScript.Sleep 500 
wShshell.Run "notepad c:\services list.txt"
WScript.Sleep 500 
end if
WScript.Quit

end if
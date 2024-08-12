On Error Resume Next
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objGetComputerList = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("outputtst.txt", True)
Set fsoReadComputerList = objGetComputerList.OpenTextFile("servers.txt", 1, TristateFalse)
aryComputers = Split(fsoReadComputerList.ReadAll, vbCrLf)
fsoReadComputerList.Close
    
For Each strComputer In aryComputers
	'strComputer = "."


Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colAdapters = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
 
n = 1
objFile.WriteLine
 
For Each objAdapter in colAdapters
   objFile.WriteLine "Network Adapter " & n & " for " & objAdapter.DNSHostName
   objFile.WriteLine "================="
   objFile.WriteLine "  Description: " & objAdapter.Description
 
   objFile.WriteLine "  Physical (MAC) address: " & objAdapter.MACAddress
   objFile.WriteLine "  Host name:              " & objAdapter.DNSHostName
 
   If Not IsNull(objAdapter.IPAddress) Then
      For i = 0 To UBound(objAdapter.IPAddress)
         objFile.WriteLine "  IP address:             " & objAdapter.IPAddress(i)
      Next
   End If
 
   If Not IsNull(objAdapter.IPSubnet) Then
      For i = 0 To UBound(objAdapter.IPSubnet)
         objFile.WriteLine "  Subnet:                 " & objAdapter.IPSubnet(i)
      Next
   End If
 
   If Not IsNull(objAdapter.DefaultIPGateway) Then
      For i = 0 To UBound(objAdapter.DefaultIPGateway)
         objFile.WriteLine "  Default gateway:        " & _
             objAdapter.DefaultIPGateway(i)
      Next
   End If
 
   objFile.WriteLine
   objFile.WriteLine "  DNS"
   objFile.WriteLine "  ---"
   objFile.WriteLine "    DNS servers in search order:"
 
   If Not IsNull(objAdapter.DNSServerSearchOrder) Then
      For i = 0 To UBound(objAdapter.DNSServerSearchOrder)
         objFile.WriteLine "      " & objAdapter.DNSServerSearchOrder(i)
      Next
   End If
 
   objFile.WriteLine "    DNS domain: " & objAdapter.DNSDomain
 
   If Not IsNull(objAdapter.DNSDomainSuffixSearchOrder) Then
      For i = 0 To UBound(objAdapter.DNSDomainSuffixSearchOrder)
         objFile.WriteLine "    DNS suffix search list: " & _
             objAdapter.DNSDomainSuffixSearchOrder(i)
      Next
   End If
 
   objFile.WriteLine
   objFile.WriteLine "  DHCP"
   objFile.WriteLine "  ----"
   objFile.WriteLine "    DHCP enabled:        " & objAdapter.DHCPEnabled
   objFile.WriteLine "    DHCP server:         " & objAdapter.DHCPServer
 
   If Not IsNull(objAdapter.DHCPLeaseObtained) Then
      utcLeaseObtained = objAdapter.DHCPLeaseObtained
      strLeaseObtained = WMIDateStringToDate(utcLeaseObtained)
   Else
      strLeaseObtained = ""
   End If
   objFile.WriteLine "    DHCP lease obtained: " & strLeaseObtained
 
   If Not IsNull(objAdapter.DHCPLeaseExpires) Then
      utcLeaseExpires = objAdapter.DHCPLeaseExpires
      strLeaseExpires = WMIDateStringToDate(utcLeaseExpires)
   Else
      strLeaseExpires = ""
   End If
   objFile.WriteLine "    DHCP lease expires:  " & strLeaseExpires
 
   objFile.WriteLine
   objFile.WriteLine "  WINS"
   objFile.WriteLine "  ----"
   objFile.WriteLine "    Primary WINS server:   " & objAdapter.WINSPrimaryServer
   objFile.WriteLine "    Secondary WINS server: " & objAdapter.WINSSecondaryServer
   objFile.WriteLine
 
   n = n + 1
 


Next
Next
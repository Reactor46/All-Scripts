sn = GetSerialNumber  
arrMonitors = GetMonitorSerials  
strDescription = "Computer: " & sn  
For intMon = LBound(arrMonitors) To UBound(arrMonitors)  
        strDescription = strDescription & " Monitor " & intMon + 1 & ": " & arrMonitors(intMon)  
Next  
'MsgBox strDescription  
UpdateDescription(strDescription)  
  
   
Function GetSerialNumber   
        strComputer = "."   
        Set objWMIService = GetObject("winmgmts:" _   
            & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")   
        Set colBIOS = objWMIService.ExecQuery _   
            ("Select * from Win32_BIOS")   
        For each objBIOS In colBIOS   
                GetSerialNumber = objBIOS.SerialNumber   
        Next   
End Function 
   
Sub UpdateDescription(strDescription)   
        Set objSysInfo = CreateObject("ADSystemInfo")
On Error Resume Next            
        Set objComputer = GetObject("LDAP://" & objSysInfo.ComputerName)   
        objComputer.Description = strDescription   
        objComputer.SetInfo   
End Sub  
  
Function GetMonitorSerials  
        Const HKEY_LOCAL_MACHINE = &H80000002  
        strKey = "SYSTEM\CurrentControlSet\Control\GraphicsDrivers\Connectivity"  
        strComputer = "."  
        strMS = ""  
        Set objRegistry = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer & "\root\default:StdRegProv")  
        If objRegistry.EnumKey(HKEY_LOCAL_MACHINE, strKey, arrKeyNames) = 0 Then  
                If IsNull(arrKeyNames) = False Then  
                        For Each strKeyName In arrKeyNames  
                                If strMS = "" Then  
                                        strMS = strKeyName  
                                Else  
                                        strMS = strMS & ";" & strKeyName  
                                End If  
                        Next  
                End If  
        End If  
        GetMonitorSerials = Split(strMS, ";")  
End Function
'
'Version Info
'$Id$

'Server Details
'==============
hostName="servicedesk"
portNo="8080"
accountName="ACCOUNT_NAME"
siteName="Site_NAME"

'********** DO NOT MODIFY ANY CODE BELOW THIS **********
protocol="http"
serviceTag=""
computerName=""
macAddress=""
'Script Mode Details
'===================
isAgentMode=false
silentMode = false
debugmode = false
agentTaskID="NO_AGENT_TASK_ID"
ae_service = "ManageEngine AssetExplorer Agent"
argCount = WScript.Arguments.Count
osVersion = ""
if(argCount>0) Then
    for i=0 to argCount-1
        if( i = 0 And StrComp(WScript.Arguments(i),"-help",1) = 0) Then
            correctUsage
            WScript.quit(0)
        End if
        if ((Not silentMode) And (Not debugmode)) Then  ' Option to run the script either in silent mode or debug mode.
        if (StrComp((WScript.Arguments(i)),"-SilentMode",1) = 0) Then
            silentMode = true
            elseif (StrComp(WScript.Arguments(i),"-debug",1) = 0) Then
                debugmode = true
            end if
        end if

        if (StrComp(WScript.Arguments(i),"-fs",1) = 0) Then
            if (i < argCount-1) And (StrComp(WScript.Arguments(i+1),"-debug",1) <> 0) And (StrComp(WScript.Arguments(i+1),"-SilentMode",1) <> 0) And (StrComp(WScript.Arguments(i+1),"-out",1) <> 0) Then
                filesearch=Wscript.Arguments(i+1)
            else
                correctUsage
                Wscript.quit(0)

            End if
        end if

        if (StrComp(WScript.Arguments(i),"-out",1) = 0) Then
            if (i < argCount-1) And (StrComp(WScript.Arguments(i+1),"-debug",1) <> 0) And (StrComp(WScript.Arguments(i+1),"-SilentMode",1) <> 0) And (StrComp(WScript.Arguments(i+1),"-fs",1) <> 0) Then
                isAgentMode = true
                agentTaskID=WScript.Arguments(i+1)

            else
                correctUsage
                WScript.quit(0)
            End if
        end if

        if (i = 0) Then
            if( (StrComp(WScript.Arguments(i),"-debug",1) <> 0) And (StrComp(WScript.Arguments(i),"-SilentMode",1) <> 0) And (StrComp(WScript.Arguments(i),"-fs",1) <> 0) And (StrComp(WScript.Arguments(i),"-out",1) <> 0))Then
                isAgentMode = true
                agentTaskID=WScript.Arguments(0)
            End if
        end if
    Next
end if

'Save Settings File Configuration
'================================
saveXMLFile=false
computerNameForFile="NO_COMPUTER_NAME"

'XML Version/Encoding Information
'================================
xmlVersion="1.0"
xmlEncoding="UTF-8"

strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set objWMIService2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\wmi")

set adobeSoftHavingLicKeys = CreateObject("Scripting.Dictionary")

const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CURRENT_USER = &H80000001
const HKEY_USERS = &H80000003
const doubleQuote=""""
const backSlash="\"
const SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056
newLineConst = VBCrLf
spaceString = " "
equalString = "="
supportMailID="assetexplorer-support@manageengine.com"
set sqlSoftList = CreateObject("Scripting.Dictionary")

initSQLSoftList
xmlInfoString = "<?xml"
xmlInfoString = addCategoryData(xmlInfoString, "version",xmlVersion)
xmlInfoString = addCategoryData(xmlInfoString, "encoding",xmlEncoding)
xmlInfoString = xmlInfoString & "?>"

outputText = xmlInfoString & newLineConst
outputText = outputText &  "<DocRoot>"

'Adding Agent Scan Key Info
'==========================
if(isAgentMode) Then
    agentTaskInfo  = "<agentTaskInfo  "
    agentTaskInfo  = addCategoryData(agentTaskInfo, "AgentTaskID",agentTaskID)
    agentTaskInfo  = agentTaskInfo & "/>"
    outputText = outputText & agentTaskInfo
end if

'Adding Script Information
'=========================
scriptVersion="5.5"
scriptVersionInfo = "<scriptVersion "
scriptVersionInfo = addCategoryData(scriptVersionInfo, "Version",scriptVersion)
scriptVersionInfo = scriptVersionInfo & "/>"
outputText = outputText & scriptVersionInfo

'Adding Scan Script Information
'=========================
scanScriptVersion="1.0.21"
scanScriptVersionInfo = "<scanScriptInfo "
scanScriptVersionInfo = addCategoryData(scanScriptVersionInfo, "Version",scanScriptVersion)
scanScriptVersionInfo = scanScriptVersionInfo & "/>"
outputText = outputText & scanScriptVersionInfo

'Data Fetching Starts
'====================
outputText = outputText & "<Hardware_Info>"
dataText = ""

'Compuer System Info
'===================
dataText = "<Computer "
getDomainName=true

'Get domain name from registry
'------------------------------
On Error Resume Next
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

is64BitOS = false
objReg.GetStringValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "PROCESSOR_ARCHITECTURE", osArch
if not isNULL(osArch) then
    pos=InStr(osArch,"64")
    if pos>0 Then
        is64BitOS = true
    End if
End if


query="Select * from Win32_ComputerSystem"
Set queryResult = objWMIService.ExecQuery (query)
For Each iterResult in queryResult
    csdata = ""
    computerNameForFile = LCase(iterResult.Caption & "")
    computerName = computerNameForFile
    'csdata = addCategoryData(csdata, "Name", computerNameForFile)
    manufacturer = iterResult.Manufacturer
    csdata = addCategoryData(csdata, "Manufacturer", manufacturer)
    csdata = addCategoryData(csdata, "PrimaryOwnerName", iterResult.PrimaryOwnerName)
    modelPartNo = iterResult.Model
    csdata = addCategoryData(csdata, "UserName", iterResult.UserName)
    csdata = addCategoryData(csdata, "WorkGroup", iterResult.WorkGroup)
    csdata = addCategoryData(csdata, "TotalPhysicalMemory", iterResult.TotalPhysicalMemory)
	csdata = addCategoryData(csdata, "LogicalCPUCount", iterResult.NumberOfLogicalProcessors)
    domainName = iterResult.Domain & ""
    if (not ISNULL(domainName) and domainName<>"" )then
    	csdata = addCategoryData(csdata, "DomainName", domainName)
    else
	    domainKeyRoot = "SYSTEM\ControlSet001\Services\Tcpip\parameters\"
	    objReg.GetStringValue HKEY_LOCAL_MACHINE, domainKeyRoot, "Domain", domainName

	    if ISNULL(domainName) then
		    domainName = ""
	    end if
	    csdata = addCategoryData(csdata, "DomainName", domainName)
    end if
    domainRole = iterResult.DomainRole
    csdata = addCategoryData(csdata, "DomainRole", domainRole)
Next

query="Select * from Win32_ComputerSystemProduct"
Set queryResult = objWMIService.ExecQuery (query)
For Each iterResult in queryResult
	model = iterResult.Name
	lenovoModel = iterResult.Version
Next

if (StrComp(manufacturer,"LENOVO",1)<>0) then
    csdata = addCategoryData(csdata, "Model", model)
else
   csdata = addCategoryData(csdata, "Model", lenovoModel&" ("&modelPartNo&")") 
End if

if ((not ISNull(domainName)) and InStr(domainName,".") > 0) then
    computerName = computerName&"."&LCase(domainName)
    computerNameForFile = computerName
End if


query="Select * from Win32_BIOS"
Set queryResult = objWMIService.ExecQuery (query)
For Each iterResult in queryResult
    biosdata = ""

    biosdata = addCategoryData(biosdata, "BiosName", iterResult.Caption)
    biosdata = addCategoryData(biosdata, "BiosManufacturer", iterResult.Manufacturer)
    biosdata = addCategoryData(biosdata, "BiosVersion", iterResult.Version)
    biosdata = addCategoryData(biosdata, "SMBiosVersion", iterResult.SMBIOSBIOSVersion)
    biosdata = addCategoryData(biosdata, "BiosDate", iterResult.ReleaseDate)
    serviceTag = iterResult.SerialNumber
    biosdata = addCategoryData(biosdata, "ServiceTag", serviceTag)
Next
'dataText = dataText & biosdata

query="select DNSDomain,DNSHostName from Win32_NetworkAdapterConfiguration where DNSDomain!=null AND DNSHostName!=null"
Set queryResult = objWMIService.ExecQuery (query)
dnsDomain= ""
dnsHostName=""
For Each iterResult in queryResult
    dnsdata = ""
    dnsDomain = iterResult.DNSDomain
    dnsHostName = iterResult.DNSHostName
    dnsdata = addCategoryData(dnsdata, "DNSDomain", dnsDomain)
    dnsdata = addCategoryData(dnsdata, "DNSHostName", dnsHostName)
Next

if((InStr(computerName,".") = 0) and (not isNull(dnsDomain)) and (InStr(dnsDomain,".")>0))then
    computerName = computerName&"."&dnsDomain
    computerNameForFile = computerName
end if
nameAttr = addCategoryData("", "Name", computerName)
dataText = dataText & nameAttr & csdata
dataText = dataText & biosdata
dataText = dataText & dnsdata

query="Select MemoryDevices from Win32_PhysicalMemoryArray where Use=3"
Set queryResult = objWMIService.ExecQuery (query)
For Each iterResult in queryResult
    mscount = ""
    mscount = addCategoryData(mscount, "MemorySlotsCount", iterResult.MemoryDevices)
Next
dataText = dataText & mscount

query="Select ChassisTypes from Win32_SystemEnclosure"
Set queryResult = objWMIService.ExecQuery (query)
For Each iterResult in queryResult
    isLaptop=""
    labtopIterator = iterResult.ChassisTypes
    for each laptop in labtopIterator
        isLaptop = laptop
    Next

Next
dataText = addCategoryData(dataText, "isLaptop", isLaptop)
dataText = dataText & "/>"
outputText = outputText & dataText
Err.clear

'Operating System Info
'=====================
On Error Resume Next
    getComputerName=true
    query="select * from Win32_OperatingSystem"
    Set queryResult = objWMIService.ExecQuery (query)
    dataText = "<OperatingSystem "
    For Each iterResult in queryResult
        osdata = ""
        osdata = addCategoryData(osdata, "Name", iterResult.Caption)
        osdata = addCategoryData(osdata, "Version", iterResult.Version)
        osdata = addCategoryData(osdata, "BuildNumber", iterResult.BuildNumber)
        osdata = addCategoryData(osdata, "ServicePackMajorVersion", iterResult.ServicePackMajorVersion)
        osdata = addCategoryData(osdata, "ServicePackMinorVersion", iterResult.ServicePackMinorVersion)
        osdata = addCategoryData(osdata, "SerialNumber", iterResult.SerialNumber)
        osdata = addCategoryData(osdata, "TotalVisibleMemorySize", iterResult.TotalVisibleMemorySize)
        osdata = addCategoryData(osdata, "FreePhysicalMemory", iterResult.FreePhysicalMemory)
        osdata = addCategoryData(osdata, "TotalVirtualMemorySize", iterResult.TotalVirtualMemorySize)
        osdata = addCategoryData(osdata, "FreeVirtualMemory", iterResult.FreeVirtualMemory)
        osArchitecture = iterResult.OSArchitecture
        if (isNULL(osArchitecture) or Trim(osArchitecture)="") then
            if(is64BitOS)then
                osArchitecture = "64-bit"
            else
                osArchitecture = "32-bit"
            end if
        end if
        osdata = addCategoryData(osdata, "OSArchitecture", osArchitecture)
        osVersion = iterResult.Version
    Next
    dataText = dataText & osdata & "/>"
    outputText = outputText & dataText
Err.clear

'CPU Info
'========

Dim procIdList()
On Error Resume Next
    query="Select * from Win32_Processor"
    Set queryResult = objWMIService.ExecQuery (query)
    count = 0
    dataText = "<CPU>"
    For Each iterResult in queryResult
        count = count+1
        dataText = dataText & "<CPU_" & count & " "
        objReg.GetStringValue HKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\CentralProcessor\0", "ProcessorNameString", cpuName
        if (not isNULL(cpuName) and cpuName<>"") then
            processorName = cpuName
        else
            processorName = iterResult.Name
        end if
        dataText = addCategoryData(dataText, "CPUName", processorName)
        dataText = addCategoryData(dataText, "CPUSpeed", iterResult.MaxClockSpeed)
        dataText = addCategoryData(dataText, "CPUStepping", iterResult.Stepping)
        dataText = addCategoryData(dataText, "CPUManufacturer", iterResult.Manufacturer)
        dataText = addCategoryData(dataText, "CPUModel", iterResult.Family)
        dataText = addCategoryData(dataText, "CPUSerialNo", iterResult.UniqueId)
        dataText = addCategoryData(dataText, "NumberOfCores", iterResult.NumberOfCores)
        dataText = dataText & "/>"
    Next
    outputText = outputText & dataText & "</CPU>"
Err.clear

'MemoryModule Info
'=================
On Error Resume Next
    query="Select * from Win32_PhysicalMemory"
    Set queryResult = objWMIService.ExecQuery (query)
    dataText = "<MemoryModule>"
    count=0
    For Each iterResult in queryResult
        count=count+1
        dataText = dataText & "<MemoryModule_" & count & " "
        dataText = addCategoryData(dataText, "Name", iterResult.Tag)
        dataText = addCategoryData(dataText, "Capacity", iterResult.Capacity)
        dataText = addCategoryData(dataText, "BankLabel", iterResult.BankLabel)
        dataText = addCategoryData(dataText, "DeviceLocator", iterResult.DeviceLocator)
        dataText = addCategoryData(dataText, "MemoryType", iterResult.MemoryType)
        dataText = addCategoryData(dataText, "TypeDetail", iterResult.TypeDetail)
        dataText = addCategoryData(dataText, "Speed", iterResult.Speed)
        dataText = dataText & "/>"
    Next
    dataText = dataText & "</MemoryModule>"
    outputText = outputText & dataText
Err.clear


'HardDisc Info
'=============
On Error Resume Next
    query="Select * from Win32_DiskDrive"
    Set queryResult = objWMIService.ExecQuery (query)
    dataText = "<HardDisk>"
    count=0
    For Each iterResult in queryResult
        count=count+1
        dataText = dataText & "<HardDisk_" & count & " "
        dataText = addCategoryData(dataText, "HDName", iterResult.Caption)
        dataText = addCategoryData(dataText, "HDModel", iterResult.Model)
        dataText = addCategoryData(dataText, "HDSize", iterResult.Size)
        HDSerialNumberId = iterResult.DeviceID
        HDSerialNumberId = Replace(HDSerialNumberId,"\","\\")
        query="Select SerialNumber from  CIM_PhysicalMedia where Tag = " & doubleQuote & HDSerialNumberId & doubleQuote
        Set queryResultForSN = objWMIService.ExecQuery (query)
        For Each iterResultSN in queryResultForSN
            serialNo = iterResultSN.SerialNumber
        Next
        if not ISNULL(serialNo) And serialNo<>"" then
            dataText = addCategoryData(dataText, "HDSerialNumber", serialNo)
        else
            dataText = addCategoryData(dataText, "HDSerialNumber", "HardDiskSerialNumber")
        End if
        dataText = addCategoryData(dataText, "HDDescription", iterResult.Description)
        dataText = addCategoryData(dataText, "HDManufacturer", iterResult.Manufacturer)
        dataText = addCategoryData(dataText, "TotalCylinders", iterResult.TotalCylinders)
        dataText = addCategoryData(dataText, "BytesPerSector", iterResult.BytesPerSector)
        dataText = addCategoryData(dataText, "SectorsPerTrack", iterResult.SectorsPerTrack)
        dataText = dataText & "/>"
    Next
    dataText = dataText & "</HardDisk>"
    outputText = outputText & dataText
Err.clear

'LogicalDisk Info
'================
On Error Resume Next
    query="Select * from Win32_LogicalDisk"
    Set queryResult = objWMIService.ExecQuery (query)
    dataText = "<LogicDrive>"
    count=0
    For Each iterResult in queryResult
        count=count+1
        dataText = dataText & "<LogicDrive_" & count & " "
        dataText = addCategoryData(dataText, "Name", iterResult.Caption)
        dataText = addCategoryData(dataText, "Description", iterResult.Description)
        discType = getDiskType(iterResult.DriveType)
        dataText = addCategoryData(dataText, "Type", discType)
        dataText = addCategoryData(dataText, "Size", iterResult.Size)
        dataText = addCategoryData(dataText, "FreeSpace", iterResult.FreeSpace)
        dataText = addCategoryData(dataText, "SerialNumber", iterResult.VolumeSerialNumber)
        dataText = addCategoryData(dataText, "FileSystem", iterResult.FileSystem)
        dataText = dataText & "/>"
    Next
    dataText = dataText & "</LogicDrive>"
    outputText = outputText & dataText
Err.clear

'PhysicalDrive Info
'==================
On Error Resume Next
    count=0
    query="Select * from CIM_MediaAccessDevice"
    Set queryResult = objWMIService.ExecQuery (query)
    dataText = "<PhysicalDrive>"
    For Each iterResult in queryResult
        count=count+1
        dataText = dataText & "<PhysicalDrive_" & count & " "
        dataText = addCategoryData(dataText, "Name", iterResult.Caption)
        dataText = addCategoryData(dataText, "Description", iterResult.Description)
        'dataText = addCategoryData(dataText, "Manufacturer", iterResult.Manufacturer)
        dataText = dataText & "/>"
    Next
    dataText = dataText & "</PhysicalDrive>"
    outputText = outputText & dataText
Err.clear



'KeyBoard Info
'=============
On Error Resume Next
    query="Select * from Win32_KeyBoard"
    Set queryResult = objWMIService.ExecQuery (query)
    dataText = ""
    For Each iterResult in queryResult
        dataText = "" 'In dll the final iteration only added...
        dataText = dataText & "<KeyBoard "
        dataText = addCategoryData(dataText, "Name", iterResult.Caption)
        dataText = dataText & "/>"
    Next
    outputText = outputText & dataText
Err.clear

'Mouse Info
'===========
On Error Resume Next
    query="Select * from Win32_PointingDevice"
    Set queryResult = objWMIService.ExecQuery (query)
    dataText = ""
    For Each iterResult in queryResult
        dataText = "" 'In dll the final iteration only added...
        dataText = dataText & "<Mouse "
        dataText = addCategoryData(dataText, "Name", iterResult.Name)
        dataText = addCategoryData(dataText, "ButtonsCount", iterResult.NumberOfButtons)
        dataText = addCategoryData(dataText, "Manufacturer", iterResult.Manufacturer)
        dataText = dataText & "/>"
    Next
    outputText = outputText & dataText
Err.clear


'Monitor Info
'============
Dim sMultiStrings()
On Error Resume Next
    query="SELECT * FROM Win32_PnPEntity where deviceid like 'DISPLAY%'"
    Set queryResult = objWMIService.ExecQuery (query)
        dataText = "<Monitor>"
        monitorCount=0
        For Each iterResult in queryResult          
            MonitorType = iterResult.Caption
            Manufacturer = iterResult.Manufacturer      
            pnpDeviceId = iterResult.PNPDeviceID
            subKey = "SYSTEM\CurrentControlSet\Enum\" & pnpDeviceId & "\Device Parameters"
            objReg.GetBinaryValue HKEY_LOCAL_MACHINE, subKey, "EDID", arrRawEDID
            If(Not isNull(arrRawEDID))Then
                productCode = getProductCode(arrRawEDID)
                manufacturerCode = getManufacturerCode(arrRawEDID)
                monitorSize = getMonitorSize(arrRawEDID)
                monitorMode = isAnalog(arrRawEDID)
                matchingArray = Split( "0 0 0 255")
                indexArray = Split( "54 72 90 108" )
                serialNumber = readValueFromEDID(matchingArray,indexArray,arrRawEDID)       
                matchingArray = Split( "0 0 0 252")
                indexArray = Split( "54 72 90 108" )
                MonitorName = readValueFromEDID(matchingArray,indexArray,arrRawEDID)        
                IF(Not isNULL(MonitorName) and MonitorName <> "") Then
                    MonitorType = MonitorName
                END IF
                If (InStr(MonitorType,monitorMode)=0 and (not isNull(monitorMode)) and monitorMode <> "") Then
                    MonitorType = MonitorType & " (" & monitorMode & ")"
                End IF
            End If
            productID = manufacturerCode & productCode
            If(not IsNull(productID) and productID <> "")Then               
                MonitorType = MonitorType & " (" & productID & ")"
            End If
            IF(NOT ISNULL(MonitorType)) Then
                monitorCount=monitorCount+1
                dataText = dataText & "<Monitor_" & monitorCount & ""
                dataText = addCategoryData(dataText, "Name", MonitorType)
                dataText = addCategoryData(dataText, "DisplayType", MonitorType)
                dataText = addCategoryData(dataText, "MonitorType", MonitorType)
                dataText = addCategoryData(dataText, "Manufacturer", Manufacturer)
                dataText = addCategoryData(dataText, "Resolution", monitorSize)
                dataText = addCategoryData(dataText, "Height", iterResult.ScreenHeight)
                dataText = addCategoryData(dataText, "Width", iterResult.ScreenWidth)
                dataText = addCategoryData(dataText, "XPixels", iterResult.PixelsPerXLogicalInch)
                dataText = addCategoryData(dataText, "YPixels", iterResult.PixelsPerYLogicalInch)
                dataText = addCategoryData(dataText, "SerialNumber", serialNumber) 
                dataText = dataText & "/>"   
            END IF    
        Next   
    dataText = dataText & "</Monitor>"
    outputText = outputText & dataText
Err.clear

'Network Info
'=============
On Error Resume Next
    query="Select * from Win32_NetworkAdapterConfiguration where IPEnabled = True"
    Set queryResult = objWMIService.ExecQuery (query)
    dataText = "<Network>"
    count=0
    For Each iterResult in queryResult
        count=count+1
        dataText = dataText & "<Network_" & count & " "
        'To Remove the NIC Index from the NIC Caption
        nwCaption = getNetworkCaption(iterResult.Caption)
        dataText = addCategoryData(dataText, "Name", nwCaption)
        dataText = addCategoryData(dataText, "Index", iterResult.Index)
        macAddress = macAddress & iterResult.MACAddress & "###"
        dataText = addCategoryData(dataText, "MACAddress", iterResult.MACAddress)
        dataText = addCategoryData(dataText, "DNSDomain", iterResult.DNSDomain)
        dataText = addCategoryData(dataText, "DNSHostName", iterResult.DNSHostName)
        dataText = addCategoryData(dataText, "DHCPEnabled", iterResult.DHCPEnabled)
        dataText = addCategoryData(dataText, "DHCPLeaseObtained", iterResult.DHCPLeaseObtained)
        dataText = addCategoryData(dataText, "DHCPLeaseExpires", iterResult.DHCPLeaseExpires)
        dataText = addCategoryData(dataText, "DHCPServer", iterResult.DHCPServer)
        ipIterator = iterResult.IPAddress
        ipAddress=""
        for each ipaddr in ipIterator
            if (ipAddress <> "") then
                ipAddress = ipAddress & "-" & ipaddr
            else
                ipAddress = ipaddr
            end if
        Next
        dataText = addCategoryData(dataText, "IpAddress", ipAddress)
        ipIterator = iterResult.DefaultIPGateway
        ipAddress=""
        for each ipaddr in ipIterator
            if (ipAddress <> "") then
                ipAddress = ipAddress & "-" & ipaddr
            else
                ipAddress = ipaddr
            end if
        Next
        dataText = addCategoryData(dataText, "Gateway", ipAddress)
        ipIterator = iterResult.DNSServerSearchOrder
        ipAddress=""
        for each ipaddr in ipIterator
            if (ipAddress <> "") then
                ipAddress = ipAddress & "-" & ipaddr
            else
                ipAddress = ipaddr
            end if
        Next
        dataText = addCategoryData(dataText, "DnsServer", ipAddress)
        ipIterator = iterResult.IPSubnet
        ipAddress=""
        for each ipaddr in ipIterator
            if (ipAddress <> "") then
                ipAddress = ipAddress & "-" & ipaddr
            else
                ipAddress = ipaddr
            end if
        Next
        dataText = addCategoryData(dataText, "Subnet", ipAddress)
        dataText = dataText & "/>"
    Next
    dataText = dataText & "</Network>"
    outputText = outputText & dataText
Err.clear

'SoundCard Info
'============
On Error Resume Next
    query="Select * from Win32_SoundDevice"
    Set queryResult = objWMIService.ExecQuery (query)
    dataText = ""
    For Each iterResult in queryResult
        dataText = "" 'In dll the final iteration only added...
        dataText = dataText & "<SoundCard "
        dataText = addCategoryData(dataText, "SoundCardName", iterResult.Caption)
        dataText = addCategoryData(dataText, "SoundCardManufacturer", iterResult.Manufacturer)
        dataText = dataText & "/>"
    Next
    outputText = outputText & dataText
Err.clear

'VideoCard Info
'==================
On Error Resume Next
    query="Select * from Win32_VideoController where Availability!=8"
    Set queryResult = objWMIService.ExecQuery (query)
    dataText = "<VideoCard>"
    count=0
    For Each iterResult in queryResult
        count=count+1
        dataText = dataText & "<VideoCard_" & count & " "
        dataText = addCategoryData(dataText, "VideoCardName", iterResult.Caption)
        dataText = addCategoryData(dataText, "VideoCardChipset", iterResult.VideoProcessor)
        dataText = addCategoryData(dataText, "VideoCardMemory", iterResult.AdapterRAM)
        dataText = dataText & "/>"
    Next

    dataText = dataText & "</VideoCard>"
    outputText = outputText & dataText
Err.clear

'SerialPort Info
'==================
On Error Resume Next
    query="Select * from Win32_SerialPort"
    Set queryResult = objWMIService.ExecQuery (query)
    dataText = "<SerialPort>"
    count=0
    For Each iterResult in queryResult
        count=count+1
        dataText = dataText & "<SerialPort_" & count & " "
        dataText = addCategoryData(dataText, "Name", iterResult.Caption)
        dataText = addCategoryData(dataText, "BaudRate", iterResult.MaxBaudRate)
        dataText = addCategoryData(dataText, "Status", iterResult.Status)
        dataText = dataText & "/>"
    Next
    dataText = dataText & "</SerialPort>"
    outputText = outputText & dataText
Err.clear

'ParallelPort Info
'==================
On Error Resume Next
    query="Select * from Win32_ParallelPort"
    Set queryResult = objWMIService.ExecQuery (query)
    dataText = "<ParallelPort>"
    count=0
    For Each iterResult in queryResult
        count=count+1
        dataText = dataText & "<ParallelPort_" & count & " "
        dataText = addCategoryData(dataText, "Name", iterResult.Caption)
        dataText = addCategoryData(dataText, "Status", iterResult.Status)
        dataText = dataText & "/>"
    Next
    dataText = dataText & "</ParallelPort>"
    outputText = outputText & dataText
Err.clear

'USB Info
'========
On Error Resume Next
    query="Select * from Win32_USBController"
    Set queryResult = objWMIService.ExecQuery (query)
    dataText = "<USB>"
    count=0
    For Each iterResult in queryResult
        count=count+1
        dataText = dataText & "<USB_" & count & " "
        dataText = addCategoryData(dataText, "Name", iterResult.Caption)
        dataText = addCategoryData(dataText, "Manufacturer", iterResult.Manufacturer)
        dataText = dataText & "/>"
    Next
    dataText = dataText & "</USB>"
    outputText = outputText & dataText
Err.clear



'Printer Info
'=================
On Error Resume Next
    set printersList = CreateObject("Scripting.Dictionary")
    dataText = "<Printer>"
    count=0
    query="Select * from Win32_Printer Where Network=False"
    Set queryResult = objWMIService.ExecQuery (query)
    For Each iterResult in queryResult
        printerName = iterResult.Caption
        if(Not printersList.Exists(printerName))then
            count=count+1
            printersList.add printerName,count
            On Error Resume Next
            dataText = dataText & "<Printer_" & count & " "
            dataText = addCategoryData(dataText, "Name", printerName)
            dataText = addCategoryData(dataText, "Model", iterResult.DriverName)
            dataText = addCategoryData(dataText, "Default", iterResult.Default)
            dataText = addCategoryData(dataText, "Network", iterResult.Network)
            dataText = addCategoryData(dataText, "Local", iterResult.Local)
            dataText = addCategoryData(dataText, "PortName", iterResult.PortName)
            dataText = addCategoryData(dataText, "Location", iterResult.Location)
            dataText = addCategoryData(dataText, "Comment", iterResult.Comment)
            dataText = addCategoryData(dataText, "ServerName", iterResult.ServerName)
            Err.clear
            dataText = dataText & "/>"
        end if
    Next

    query="Select * from Win32_Printer Where Network=True"
    Set queryResult = objWMIService.ExecQuery (query)
    For Each iterResult in queryResult
        printerName = iterResult.Caption
        if(Not printersList.Exists(printerName))then
            count=count+1
            printersList.add printerName,count
            On Error Resume Next
            dataText = dataText & "<Printer_" & count & " "
            dataText = addCategoryData(dataText, "Name", printerName)
            dataText = addCategoryData(dataText, "Model", iterResult.DriverName)
            dataText = addCategoryData(dataText, "Default", iterResult.Default)
            dataText = addCategoryData(dataText, "Network", iterResult.Network)
            dataText = addCategoryData(dataText, "Local", iterResult.Local)
            dataText = addCategoryData(dataText, "PortName", iterResult.PortName)
            dataText = addCategoryData(dataText, "Location", iterResult.Location)
            dataText = addCategoryData(dataText, "Comment", iterResult.Comment)
            dataText = addCategoryData(dataText, "ServerName", iterResult.ServerName)
            Err.clear
            dataText = dataText & "/>"
        end if
    Next
    if(queryResult.Count = 0) then
        dataText = getPrinterInfo(dataText,count)
    End if
    dataText = dataText & "</Printer>"
    outputText = outputText & dataText
Err.clear

'HotFix Info
'===========
On Error Resume Next
    if(isVista)then
        query="Select * from Win32_QuickFixEngineering where caption!=''" ' where condition added to avoid some unknown hotfix.(i.e) In vista machine, it results a key but could not get actual hotfix name.        
    elseif(isVistaAndAbove)then
        query="Select * from Win32_QuickFixEngineering"         
    else
        query="Select * from Win32_QuickFixEngineering where FixComments!=''"
       
    end if
    Set queryResult = objWMIService.ExecQuery (query)
        dataText = "<HotFix>"
        count=0
        For Each iterResult in queryResult
            count=count+1
            dataText = dataText & "<HotFix_" & count & " "
            dataText = addCategoryData(dataText, "HotFixID", iterResult.HotFixID)
            dataText = addCategoryData(dataText, "InstalledBy", iterResult.InstalledBy)
            dataText = addCategoryData(dataText, "InstalledOn", iterResult.InstalledOn)
            dataText = addCategoryData(dataText, "Description", iterResult.Description)
            dataText = dataText & "/>"
        Next
        dataText = dataText & "</HotFix>"
    outputText = outputText & dataText
Err.clear

'Users account details
'=====================
'do not fetch users from AD
if(not(domainRole=4 or domainRole=5))then ' if not AD
On Error Resume Next
    query="Select PartComponent from Win32_SystemUsers"
    Set queryResult = objWMIService.ExecQuery (query)
    dataText = "<UsersAccount>"
    count=0
    For Each iterResult in queryResult      
        partComponent = iterResult.PartComponent
        userNameCri = Split(Split(partComponent,"UserAccount.")(1),",")(0)
        domainNameCri = Split(Split(partComponent,"UserAccount.")(1),",")(1)
        if(ISNULL(userNameCri) or ISNULL(domainNameCri)) then
            Set accountDetails = objWMIService.Get(partComponent)
            count=count+1
            dataText = dataText & "<UsersAccount_" & count & " "
            dataText = addCategoryData(dataText, "Name", accountDetailsForUA.Name)
            dataText = addCategoryData(dataText, "Domain", accountDetailsForUA.Domain)
            dataText = addCategoryData(dataText, "FullName", accountDetailsForUA.FullName)
            dataText = addCategoryData(dataText, "Description", accountDetailsForUA.Description)
            dataText = addCategoryData(dataText, "LocalAccount", accountDetailsForUA.LocalAccount)
            dataText = addCategoryData(dataText, "Status", accountDetailsForUA.Status)
            dataText = addCategoryData(dataText, "SID", accountDetailsForUA.SID)
            dataText = dataText & "/>"
        else
            UAQuery = "Select * from  Win32_UserAccount where " & userNameCri & " AND "& domainNameCri
            UAQuery = replace(UAQuery,"""","'")
            Set accountDetailsList =objWMIService.ExecQuery (UAQuery)
            For Each accountDetailsForUA in accountDetailsList
                count=count+1
                dataText = dataText & "<UsersAccount_" & count & " "
                dataText = addCategoryData(dataText, "Name", accountDetailsForUA.Name)
                dataText = addCategoryData(dataText, "Domain", accountDetailsForUA.Domain)
                dataText = addCategoryData(dataText, "FullName", accountDetailsForUA.FullName)
                dataText = addCategoryData(dataText, "Description", accountDetailsForUA.Description)
                dataText = addCategoryData(dataText, "LocalAccount", accountDetailsForUA.LocalAccount)
                dataText = addCategoryData(dataText, "Status", accountDetailsForUA.Status)
                dataText = addCategoryData(dataText, "SID", accountDetailsForUA.SID)
                dataText = dataText & "/>"
            Next
        end if  
    Next
    dataText = dataText & "</UsersAccount>"
    outputText = outputText & dataText
Err.clear
end If
outputText = outputText & "</Hardware_Info>"

'Microsoft Keys
'=================
On Error Resume Next
outputText = outputText & "<Software_Info>"
set licenseKeys = CreateObject("Scripting.Dictionary")

licenseDataText = licenseDataText & "<MicrosoftOfficeKeys>"
count=0
strKeyPath = "SOFTWARE\Microsoft\Office"
for itr=1 to 2
    continue = true
    Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
    if(itr=2) then
        objCtx.Add "__ProviderArchitecture", 32
    elseif (is64BitOS) then
        objCtx.Add "__ProviderArchitecture", 64
    else
        continue = false
    end if
    objCtx.Add "__RequiredArchitecture", TRUE
    if(continue)then
        Set objLocator = CreateObject("Wbemscripting.SWbemLocator")
        Set objServices = objLocator.ConnectServer("localhost","root\default","","",,,,objCtx)
        Set objReg1 = objServices.Get("StdRegProv")
        objReg1.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
        If NOT ISNULL(arrSubKeys) then
            For Each subkey In arrSubKeys
                subkeyPath = strKeyPath & "\" & subkey & "\Registration"
                objReg1.EnumKey HKEY_LOCAL_MACHINE, subkeyPath, productSubKeys
                If NOT ISNULL(productSubKeys) then
                    For Each productsubkey In productSubKeys
                        count=count+1
                        licenseDataText = licenseDataText & "<Key_" & count & " "
                        licenseDataText = addCategoryData(licenseDataText, "Key", productsubkey)
                        productsubkeyPath = subkeyPath & "\" &  productsubkey
                        objReg1.GetStringValue HKEY_LOCAL_MACHINE, productsubkeyPath, "ProductID", productId
                        objReg1.GetBinaryValue HKEY_LOCAL_MACHINE, productsubkeyPath, "DigitalProductID", productKey
                        objReg1.GetStringValue HKEY_LOCAL_MACHINE, productsubkeyPath,"ConvertToEdition",convertToEdition
                        if NOT ISNULL(productId) then
                            licenseDataText = addCategoryData(licenseDataText, "ProductID", productId)
                        end if
                        if (NOT ISNULL(productKey) And NOT ISNULL(convertToEdition) And (InStr(Lcase(convertToEdition),"2010") = 0) And (InStr(Lcase(convertToEdition),"2013") = 0))then
                            key = getLicenceKey(productKey,subkey)
                            if (NOT ISNULL(key) and key<>"") then
                                licenseDataText = addCategoryData(licenseDataText, "ProductKey", key)
                                licenseKeys.add productsubkey,key
                            end if
                        end if
                        licenseDataText = licenseDataText &  " />"
                    Next
                end if
            Next
        end if
    end if
Next
licenseDataText = licenseDataText & "</MicrosoftOfficeKeys>"
Err.clear

'Windows Key
'===========
On Error Resume Next
licenseDataText = licenseDataText & "<WindowsKey "
windowsKeyRoot = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\"
count=0
strKeyPath = "SOFTWARE\Microsoft\Office"
isProductIdDetected = false
for itr=1 to 2
	continue = true
	Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
	if(itr=2) then
		objCtx.Add "__ProviderArchitecture", 32
	elseif (is64BitOS) then
		objCtx.Add "__ProviderArchitecture", 64
	else
		continue = false
	end if
	objCtx.Add "__RequiredArchitecture", TRUE
	if(continue And isProductIdDetected=false)then
		Set objLocator = CreateObject("Wbemscripting.SWbemLocator")
		Set objServices = objLocator.ConnectServer("localhost","root\default","","",,,,objCtx)
		Set objReg1 = objServices.Get("StdRegProv") 
		objReg1.GetStringValue HKEY_LOCAL_MACHINE, windowsKeyRoot, "ProductID", windowsId
		if not ISNULL(windowsId) then
			licenseDataText = addCategoryData(licenseDataText, "ProductID", windowsId)
			isProductIdDetected = true
		end if
		objReg1.GetBinaryValue HKEY_LOCAL_MACHINE, windowsKeyRoot, "DigitalProductID", windowsKeyData
        objReg1.GetStringValue HKEY_LOCAL_MACHINE, windowsKeyRoot, "CurrentVersion", CurrentVersion
        currentVerionList = split(CurrentVersion,".")
        majorVersion = CurrentVerionList(0)
        if majorVersion > 4 then
            if not ISNULL(windowsKeyData)then
                windowsKey = getLicenceKey(windowsKeyData,null)
            end if
        else
            if not ISNULL(windowsKeyData)then
                windowsKey = ConvertToKey(windowsKeyData)                
            end if
        end if
        if(InStr(Lcase(windowsKey),"bbbbb") > 0) then
            query = "Select * from softwarelicensingproduct where partialproductkey != NULL and description like '%operating system%'"
            Set queryResult = objWMIService.ExecQuery (query)
            count=0
            For Each iterResult in queryResult              
                partialproductkey = iterResult.PartialProductKey
            Next            
            if not ISNULL(partialproductkey) then
                windowsKey = "XXXXX-XXXXX-XXXXX-XXXXX-" & partialproductkey
                licenseDataText = addCategoryData(licenseDataText, "ProductKey", windowsKey)                
            end if
        else 
            if not ISNULL(windowsKeyData) then
                licenseDataText = addCategoryData(licenseDataText, "ProductKey", windowsKey)
            end if
        end if
	end if
Next
licenseDataText = licenseDataText & "/>"
Err.clear
outputText = outputText & licenseDataText
'SoftwareList Info
'=================
On Error Resume Next
classKeyPath="Software\Classes\Installer\Products"
softwareDataText = softwareDataText & "<InstalledProgramsList>"
strComputer = "."
set softList = CreateObject("Scripting.Dictionary")

Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
count=0

set autodeskSoftLicKeys = CreateObject("Scripting.Dictionary")
autodeskProductSuiteSerialNumber = ""
setAutodeskLicenses objReg

'Softwares installed under different users
'=========================================
objReg.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList",profileSIDs

If NOT ISNULL(profileSIDs) then
    For Each profileSID In profileSIDs
        objReg.EnumKey HKEY_USERS, profileSID & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",arrSubKeys
        If NOT ISNULL(arrSubKeys) then
			For Each subkey In arrSubKeys				
				subkeyPath =  profileSID & "\" & strKeyPath & "\" & subkey
				objReg1.GetDWORDValue HKEY_USERS, subkeyPath, "SystemComponent", SystemComponent
				objReg1.GetDWORDValue HKEY_USERS, subkeyPath, "WindowsInstaller", WindowsInstaller
                objReg1.GetStringValue HKEY_USERS, subkeyPath, "ParentKeyName", parentKeyName
                objReg1.GetStringValue HKEY_USERS, subkeyPath, "DisplayVersion", softwareVersion
                objReg1.GetStringValue HKEY_USERS, subkeyPath, "Publisher", softwarePublisher
                objReg1.GetStringValue HKEY_USERS, subkeyPath, "InstallLocation", softwareLocation
                objReg1.GetStringValue HKEY_USERS, subkeyPath, "InstallDate", softwareInstallDate
				objReg1.GetStringValue HKEY_USERS, subkeyPath, "ReleaseType", ReleaseType
				objReg1.GetStringValue HKEY_USERS, subkeyPath, "UninstallString", UninstallString
				objReg1.GetStringValue HKEY_USERS, subkeyPath, "DisplayName", softwareName
				keyForSoftwareUsage = profileSID & "\SOFTWARE\Microsoft\Windows\CurrentVersion\App Management\ARPCache\" & subkey
                objReg1.GetBinaryValue HKEY_USERS, keyForSoftwareUsage, "SlowInfoCache", usageData
				swUsage = getSoftwareUsage(usageData)
					if swUsage = "Not Known" Then
						objReg1.GetBinaryValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Management\ARPCache\" & subkey, "SlowInfoCache", usageData
						swUsage = getSoftwareUsage(usageData)
				End if                
				
				IF(Not ISNULL(softwareName)) THEN
					isAddSoftware = false
						IF((ISNULL(SystemComponent) or SystemComponent <> "1")) THEN
							if ((InStr(Lcase(softwareName),"adobe") > 0) or (InStr(Lcase(softwareName),"acrobat") > 0)) then
								licenseKey = getAdobeLicenseKey(objReg,subkeyPath,softwareName)
								if (licenseKey="") then
									tmpCheck = adobeSoftHavingLicKeys.Item(softwareName)
									if((Not isNULL(tmpCheck)) and (tmpCheck<>""))then
										isAddSoftware = false
									End if
								end if
								if ((licenseKey<> "") and (Not adobeSoftHavingLicKeys.Exists(softwareName)))then
									adobeSoftHavingLicKeys.add softwareName,licenseKey
								end if
							elseif(autodeskSoftLicKeys.Exists(softwareName)) then
								licenseKey = autodeskSoftLicKeys.Item(softwareName)
							elseif ((InStr(Lcase(softwareName),"autocad") > 0) or (InStr(Lcase(softwareName),"autodesk") > 0)) then
								if (InStr(Lcase(softwareName),"suite")) > 0 Then 
									licenseKey = autodeskProductSuiteSerialNumber
								End If
							elseif (isSQL(softwareName))then
								softList.add softwareName,count
								softwareName = updateSQLEdition(softwareName,HKEY_USERS,profileSID&"\SOFTWARE\Microsoft",objReg)
							End if
							IF((ISNULL(WindowsInstaller) or WindowsInstaller <> "1")) THEN
								Set re = New RegExp
								With re
									.Pattern    = "KB[0-9]{6}$"
									.IgnoreCase = False
									.Global     = False
								End With
								Set re = Nothing
								IF(NOT re.Test(subkey)) THEN
									IF(ISNULL(parentKeyName)) THEN
										IF((ISNULL(ReleaseType)) OR ((ReleaseType <> "Security Update") AND (ReleaseType <> "Update Rollup") AND (ReleaseType <> "Hotfix"))) THEN
											IF(Not ISNULL(UninstallString)) THEN
												isAddSoftware = true
											END IF
										END IF
									END IF
								END IF						
							ELSEIF(WindowsInstaller = "1") THEN
									tempSubKey = GetInstallerKeyNameFromGuid(subkey)
									tempSubKeyPath=profileSID & classKeyPath & "\" & tempSubKey
									objReg1.GetStringValue HKEY_USERS, tempSubKeyPath, "ProductName", softwareName
									isAddSoftware = true
							END IF
							
						END IF
						IF ((isAddSoftware = true) AND (Not ISNULL(softwareName)) AND (softwareName <> "")) THEN
								count=count+1
								softList.add softwareName,count
								softwareDataText = softwareDataText & "<Software_" & count & " "
								softwareDataText = addCategoryData(softwareDataText, "Name", softwareName)
								softwareDataText = addCategoryData(softwareDataText, "Version", softwareVersion)
								softwareDataText = addCategoryData(softwareDataText, "Vendor", softwarePublisher)
								softwareDataText = addCategoryData(softwareDataText, "Location", softwareLocation)
								softwareDataText = addCategoryData(softwareDataText, "InstallDate", softwareInstallDate)
								softwareDataText = addCategoryData(softwareDataText, "Usage", swUsage)
								if(licenseKey<>"" and StrComp(licenseKey,"-")<>0)then
									softwareDataText = addCategoryData(softwareDataText, "ProductKey", licenseKey)
								end if
								softwareDataText = addCategoryData(softwareDataText, "Key", subkey)
								softwareDataText = softwareDataText &  "/>"
						END IF
				END IF			
			Next
        end if
		objReg.EnumKey HKEY_USERS, profileSID&"\Software\Microsoft\Installer\Products",cuInstallerKey
		 If NOT ISNULL(cuInstallerKey) then
		 	For Each cuProductGuid In cuInstallerKey
				productFound = false
				objReg.EnumKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Installer\UserData",userDataKey
				For Each userDataKeyName In userDataKey
					IF userDataKeyName <> "S-1-5-18" Then
						objReg.EnumKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Installer\UserData\" & userDataKeyName & "\Products",lmProductGuids
						For Each lmProductGuid In lmProductGuids
							If lmProductGuid=cuProductGuid Then								
								userDataProgramKey="Software\Microsoft\Windows\CurrentVersion\Installer\UserData\" & userDataKeyName & "\Products\" & lmProductGuid & "\InstallProperties"
								objReg.GetDWORDValue HKEY_LOCAL_MACHINE, userDataProgramKey, "SystemComponent", SystemComponent
								IF((ISNULL(SystemComponent) or SystemComponent <> "1")) THEN
									objReg.GetStringValue HKEY_USERS, profileSID & "\Software\Microsoft\Installer\Products\" & cuProductGuid, "ProductName", softwareName
										objReg.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Management\ARPCache\", subKeysList
										If NOT ISNULL(subKeysList) then
											For Each subkeytemp In subKeysList
												tempSubKeyList = GetInstallerKeyNameFromGuid(subkeytemp)
												if tempSubKeyList = lmProductGuid Then
													objReg.GetBinaryValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Management\ARPCache\" & subkeytemp, "SlowInfoCache", usageData
													swUsage = getSoftwareUsage(usageData)					
												End IF
											Next
										End IF
										if swUsage = "Not Known" Then
											objReg.EnumKey HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Management\ARPCache\", subKeysList
											If NOT ISNULL(subKeysLists) then
												For Each subkeytemp In subKeysList
													tempSubKeyList = GetInstallerKeyNameFromGuid(subkeytemp)
													if tempSubKeyList = lmProductGuid Then
														objReg.GetBinaryValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Management\ARPCache\" & subkeytemp, "SlowInfoCache", usageData
														swUsage = getSoftwareUsage(usageData)
													End IF
												Next
											End IF
										End if
									
									objReg.GetStringValue HKEY_LOCAL_MACHINE, userDataProgramKey, "DisplayVersion", softwareVersion
									objReg.GetStringValue HKEY_LOCAL_MACHINE, userDataProgramKey, "Publisher", softwarePublisher
									objReg.GetStringValue HKEY_LOCAL_MACHINE, userDataProgramKey, "InstallLocation", softwareLocation
									objReg.GetStringValue HKEY_LOCAL_MACHINE, userDataProgramKey, "InstallDate", softwareInstallDate
									objReg.GetStringValue HKEY_LOCAL_MACHINE, userDataProgramKey, "ReleaseType", ReleaseType
									objReg.GetStringValue HKEY_LOCAL_MACHINE, userDataProgramKey, "UninstallString", UninstallString
                                    IF (softwareName <> "" AND (Not ISNULL(softwareName))) THEN
    									count=count+1
    									softList.add softwareName,count
    									softwareDataText = softwareDataText & "<Software_" & count & " "
    									softwareDataText = addCategoryData(softwareDataText, "Name", softwareName)
    									softwareDataText = addCategoryData(softwareDataText, "Version", softwareVersion)
    									softwareDataText = addCategoryData(softwareDataText, "Vendor", softwarePublisher)
    									softwareDataText = addCategoryData(softwareDataText, "Location", softwareLocation)
    									softwareDataText = addCategoryData(softwareDataText, "InstallDate", softwareInstallDate)
    									softwareDataText = addCategoryData(softwareDataText, "Usage", swUsage)
    									softwareDataText = addCategoryData(softwareDataText, "Key", lmProductGuid)
    									softwareDataText = softwareDataText &  "/>"
    									productFound = true
                                    END IF
								END IF
								Exit For
							End If
						Next
						If productFound = true Then
                                Exit For
                        End If
					End IF
				Next
			Next
		 End if
    Next
End if

'Softwares installed
'===================
is64BitOS = false
objReg.GetStringValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "PROCESSOR_ARCHITECTURE", osArch
if not isNULL(osArch) then
    pos=InStr(osArch,"64")
    if pos>0 Then
        is64BitOS = true
    End if
End if

for itr=1 to 2
    continue = true
    Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
    if(itr=2) then
        objCtx.Add "__ProviderArchitecture", 32
    elseif (is64BitOS) then
        objCtx.Add "__ProviderArchitecture", 64
    else
        continue = false
    end if
    objCtx.Add "__RequiredArchitecture", TRUE
    if(continue)then
        Set objLocator = CreateObject("Wbemscripting.SWbemLocator")
        Set objServices = objLocator.ConnectServer("localhost","root\default","","",,,,objCtx)
        Set objReg1 = objServices.Get("StdRegProv")
        objReg1.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
        If NOT ISNULL(arrSubKeys) then
			For Each subkey In arrSubKeys	
			   subkeyPath = strKeyPath & "\" & subkey
					objReg1.GetDWORDValue HKEY_LOCAL_MACHINE, subkeyPath, "SystemComponent", SystemComponent
					objReg1.GetDWORDValue HKEY_LOCAL_MACHINE, subkeyPath, "WindowsInstaller", WindowsInstaller
					objReg1.GetStringValue HKEY_LOCAL_MACHINE, subkeyPath, "ParentKeyName", parentKeyName
					objReg1.GetStringValue HKEY_LOCAL_MACHINE, subkeyPath, "DisplayVersion", softwareVersion
					objReg1.GetStringValue HKEY_LOCAL_MACHINE, subkeyPath, "Publisher", softwarePublisher
					objReg1.GetStringValue HKEY_LOCAL_MACHINE, subkeyPath, "InstallLocation", softwareLocation
					objReg1.GetStringValue HKEY_LOCAL_MACHINE, subkeyPath, "InstallDate", softwareInstallDate
					objReg1.GetStringValue HKEY_LOCAL_MACHINE, subkeyPath, "ReleaseType", ReleaseType
					objReg1.GetStringValue HKEY_LOCAL_MACHINE, subkeyPath, "UninstallString", UninstallString
					objReg1.GetStringValue HKEY_LOCAL_MACHINE, subkeyPath, "DisplayName", softwareName
					keyForSoftwareUsage = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Management\ARPCache\" & subkey
					objReg.GetBinaryValue HKEY_LOCAL_MACHINE, keyForSoftwareUsage, "SlowInfoCache", usageData
					swUsage = getSoftwareUsage(usageData)
					if swUsage = "Not Known" Then
						objReg.GetBinaryValue HKEY_CURRENT_USER, keyForSoftwareUsage, "SlowInfoCache", usageData
						swUsage = getSoftwareUsage(usageData)
					End if											
					IF(Not ISNULL(softwareName)) THEN
						isAddSoftware = false
						IF((ISNULL(SystemComponent) or SystemComponent <> "1")) THEN
							if ((InStr(Lcase(softwareName),"adobe") > 0) or (InStr(Lcase(softwareName),"acrobat") > 0)) then
								licenseKey = getAdobeLicenseKey(objReg,subkeyPath,softwareName)
								if (licenseKey="") then
									tmpCheck = adobeSoftHavingLicKeys.Item(softwareName)
									if((Not isNULL(tmpCheck)) and (tmpCheck<>""))then
										isAddSoftware = false
									End if
								end if
								if ((licenseKey<> "") and (Not adobeSoftHavingLicKeys.Exists(softwareName)))then
									adobeSoftHavingLicKeys.add softwareName,licenseKey
								end if
							elseif(autodeskSoftLicKeys.Exists(softwareName)) then
								licenseKey = autodeskSoftLicKeys.Item(softwareName)
							elseif ((InStr(Lcase(softwareName),"autocad") > 0) or (InStr(Lcase(softwareName),"autodesk") > 0)) then
								if (InStr(Lcase(softwareName),"suite")) > 0 Then 
									licenseKey = autodeskProductSuiteSerialNumber
								End If
							elseif (isSQL(softwareName))then
								softList.add softwareName,count
								softwareName = updateSQLEdition(softwareName,HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft",objReg1)
							End if
								IF((ISNULL(WindowsInstaller) or WindowsInstaller <> "1")) THEN
									Set re = New RegExp
									With re
										.Pattern    = "KB[0-9]{6}$"
										.IgnoreCase = False
										.Global     = False
									End With
									Set re = Nothing
									IF(NOT re.Test(subkey)) THEN
										IF(ISNULL(parentKeyName)) THEN
											IF((ISNULL(ReleaseType)) OR ((ReleaseType <> "Security Update") AND (ReleaseType <> "Update Rollup") AND (ReleaseType <> "Hotfix"))) THEN
												IF(Not ISNULL(UninstallString)) THEN
													isAddSoftware = true
												END IF
											END IF
										END IF
									END IF						
								ELSEIF(WindowsInstaller = "1") THEN
										tempSubKey = GetInstallerKeyNameFromGuid(subkey)
										tempSubKeyPath=classKeyPath & "\" & tempSubKey
										objReg.GetStringValue HKEY_LOCAL_MACHINE, tempSubKeyPath, "ProductName", softwareName
										isAddSoftware = true
								END IF
						END IF
							IF ((isAddSoftware = true) AND (Not ISNULL(softwareName)) AND (softwareName <> "")) THEN
									count=count+1
									softList.add softwareName,count
									softwareDataText = softwareDataText & "<Software_" & count & " "
									softwareDataText = addCategoryData(softwareDataText, "Name", softwareName)
									softwareDataText = addCategoryData(softwareDataText, "Version", softwareVersion)
									softwareDataText = addCategoryData(softwareDataText, "Vendor", softwarePublisher)
									softwareDataText = addCategoryData(softwareDataText, "Location", softwareLocation)
									softwareDataText = addCategoryData(softwareDataText, "InstallDate", softwareInstallDate)
									softwareDataText = addCategoryData(softwareDataText, "Usage", swUsage)
									if(licenseKey<>"" and StrComp(licenseKey,"-")<>0)then
										softwareDataText = addCategoryData(softwareDataText, "ProductKey", licenseKey)
									end if
									softwareDataText = addCategoryData(softwareDataText, "Key", subkey)
									softwareDataText = softwareDataText &  "/>"
							END IF
					END IF
			NEXT
        end if
    end if
Next

Err.clear

Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
objCtx.Add "__ProviderArchitecture", 32
objCtx.Add "__RequiredArchitecture", TRUE
Set objLocator = CreateObject("Wbemscripting.SWbemLocator")
Set objServices = objLocator.ConnectServer("localhost","root\default","","",,,,objCtx)
Set objReg1 = objServices.Get("StdRegProv")

'Enumerate IE in vista and later OS
'==================================

objReg.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Internet Explorer\", "Version", ieVersion
if not isNULL(ieVersion) then
    idx = InStr(ieVersion,".")
    If(idx>0) Then
        ieMajorVersion = Trim(Left(ieVersion,idx-1))
        if(ieMajorVersion > 8) Then
			objReg.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Internet Explorer\", "svcVersion", ieNewVersion
			if not isNULL(ieNewVersion) then
				idx = InStr(ieNewVersion,".")
				If(idx>0) Then
					ieMajorVersion = Trim(Left(ieNewVersion,idx-1))
					ieVersion = ieNewVersion
				End If
			End If
		End If        
        softwareName = "Windows Internet Explorer "&ieMajorVersion
    End If
    if(softwareName <> "" and (not softList.Exists(softwareName))) then
        count=count+1
        softList.add softwareName,count
        softwareDataText = softwareDataText & "<Software_" & count & " "
        softwareDataText = addCategoryData(softwareDataText, "Name", softwareName)
        softwareDataText = addCategoryData(softwareDataText, "Version", ieVersion)
        softwareDataText = softwareDataText &  "/>"
    End if
End  if

'file search
'===========
On Error Resume Next
if(not isNULL(filesearch) and filesearch <> "") then

    ext = stringTokenizer(filesearch,",")
    extCount = UBound(ext)

    for i=0 to extCount-1

        query="Select * from CIM_DataFile where Extension = '"&ext(i) & "'"

        Set queryResult = objWMIService.ExecQuery (query)

        For Each iterResult in queryResult
            softwareName = Mid(iterResult.Name,InStrRev(iterResult.Name,"\")+1)
            softwareVersion = iterResult.Version
            softwarePublisher = iterResult.Manufacturer
            softwareLocation = iterResult.Drive & iterResult.Path
            softwareInstallDate = iterResult.InstallDate

            If NOT ISNULL(softwareName) then
                if(softwareName <> "" and (not softList.Exists(softwareName))) then

                    count=count+1
                    softList.add softwareName,count
                    softwareDataText = softwareDataText & "<Software_" & count & " "
                    softwareDataText = addCategoryData(softwareDataText, "Name", softwareName)
                    softwareDataText = addCategoryData(softwareDataText, "Version", softwareVersion)
                    softwareDataText = addCategoryData(softwareDataText, "Vendor", softwarePublisher)
                    softwareDataText = addCategoryData(softwareDataText, "Location", softwareLocation)
                    'softwareDataText = addCategoryData(softwareDataText, "InstallDate", softwareInstallDate)
                    'softwareDataText = addCategoryData(softwareDataText, "Usage", swUsage)
                    'softwareDataText = addCategoryData(softwareDataText, "Key", subkey)
                    softwareDataText = softwareDataText &  "/>"
                end if
            end if
        Next

    Next
End if
softwareDataText = softwareDataText & "</InstalledProgramsList>"
Err.clear

'Oracle Info
'===========
On Error Resume Next
softwareDataText = softwareDataText & "<OracleInfo>"
oracleKeyRoot = "SOFTWARE\ORACLE"
objReg.EnumKey HKEY_LOCAL_MACHINE, oracleKeyRoot, arrSubKeys
count=0
If NOT ISNULL(arrSubKeys) then
    For Each subkey In arrSubKeys
        if(Left(subkey,4)="HOME" Or Left(subkey,7)="KEY_Ora") then

            subkeyPath = oracleKeyRoot & "\" & subkey
            objReg.GetStringValue HKEY_LOCAL_MACHINE, subkeyPath, "SQLPATH", sqlPath
            objReg.GetStringValue HKEY_LOCAL_MACHINE, subkeyPath, "ORACLE_GROUP_NAME", groupName
            objReg.GetStringValue HKEY_LOCAL_MACHINE, subkeyPath, "ORACLE_HOME_NAME", homeName
            objReg.GetStringValue HKEY_LOCAL_MACHINE, subkeyPath, "ORACLE_BUNDLE_NAME", bundleName
            objReg.GetStringValue HKEY_LOCAL_MACHINE, subkeyPath, "VERSION", version
            objReg.GetStringValue HKEY_LOCAL_MACHINE, subkeyPath, "ORACLE_HOME", home
            count=count+1
            softwareDataText = softwareDataText & "<Software_" & count & " "
            softwareDataText = addCategoryData(softwareDataText, "sqlPath", sqlPath)
            softwareDataText = addCategoryData(softwareDataText, "groupName", groupName)
            softwareDataText = addCategoryData(softwareDataText, "homeName", homeName)
            softwareDataText = addCategoryData(softwareDataText, "bundleName", bundleName)
            softwareDataText = addCategoryData(softwareDataText, "version", version)
            softwareDataText = addCategoryData(softwareDataText, "home", home)
            softwareDataText = addCategoryData(softwareDataText, "Key", subkey)
            softwareDataText = softwareDataText &  "/>"
        end if
    Next
end if
softwareDataText = softwareDataText & "</OracleInfo>"
softwareDataText = softwareDataText & "</Software_Info>"
outputText = outputText  & softwareDataText

'Hyper-V Details
'===============
Err.clear
isHyperV = false 
Set objWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set hypWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\virtualization")
if(Err)then
	printLog "Error while connect root\virtualization - "&getErrorMessage(Err)
else
	query="Select Caption from Msvm_ComputerSystem"
	Set queryResult = hypWMIService.ExecQuery (query)
	rowsCount = queryResult.Count
	if(Err)then
		printLog "Error after execute the query of class Msvm_ComputerSystem - "&getErrorMessage(Err)
	else
		isHyperV = true
	end if
end if

if(not isHyperV) then
	printLog "going to connect root\virtualization\v2"
	Set hypWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\virtualization\v2")
	If(Err)then
		printLog "Error while connect root\virtualization\v2 - "&getErrorMessage(Err)
	else
		query="Select Caption from Msvm_ComputerSystem"
		Set queryResult = hypWMIService.ExecQuery (query)
		printLog "Error after query 2- "&getErrorMessage(Err)
		rowsCount = queryResult.Count
		if(Err)then
			printLog "Error after execute the query of class Msvm_ComputerSystem - "&getErrorMessage(Err)
		else
			isHyperV = true
		end if
	end if
end if
printLog "Is HyperV? - "&isHyperV
if(isHyperV)then
	set objXMLDoc = CreateObject("Microsoft.xmldom") 
	printLog "error "&Err.Description
    objXMLDoc.async = False 
	
	if(not isNull(hypWMIService)) then
		query="Select Caption,Name,ElementName from Msvm_ComputerSystem where Caption!='Hosting Computer System'"
		Set queryResult = hypWMIService.ExecQuery (query)
		rowsCount = queryResult.Count
		if(Err)then
			printLog "Error after execute the query of class Msvm_ComputerSystem - "&getErrorMessage(Err)
			isHyperV = False
		Else
			vmData = "<HyperV><VirtualMachines>"
			For Each iterResult in queryResult
				vmData = vmData&"<VirtualMachine "
				vmData = addCategoryData(vmData, "VMElementId", iterResult.Name)
				vmData = addCategoryData(vmData, "VMName", iterResult.ElementName)
				elementName = iterResult.ElementName
				procQuery = "select instanceId,Limit,Reservation,Weight from MsVM_ProcessorSettingData where instanceId like '%"&iterResult.Name&"%'"
				printLog procQuery
				Set procResult = hypWMIService.ExecQuery (procQuery)
				printLog "Error after execute the query of class MsVM_ProcessorSettingData - "&Err.Description
				For Each vmProc In procResult
					vmData = addCategoryData(vmData, "CPULimitInPercentage", (vmProc.Limit/1000))
					vmData = addCategoryData(vmData, "CPUReservationInPercentage", (vmProc.Reservation/1000))
					vmData = addCategoryData(vmData, "CPUShares", vmProc.Weight)
				Next
				memQuery = "select Limit,VirtualQuantity,DynamicMemoryEnabled,Weight from MSVM_MemorySettingData where instanceId like "&doubleQuote&"%"&iterResult.Name&"%"&doubleQuote
				Set memResult = hypWMIService.ExecQuery (memQuery)
				For Each vmMem In memResult
					vmData = addCategoryData(vmData, "MemoryReservation", vmMem.VirtualQuantity)
					if (vmMem.DynamicMemoryEnabled = True) Then
						vmData = addCategoryData(vmData, "MemoryLimit", vmMem.Limit)
					End if
					vmData = addCategoryData(vmData, "MemoryShares", vmMem.Weight)
				Next
				
				Set kvpExchangeComponents = iterResult.Associators_("Msvm_SystemDevice", "Msvm_KvpExchangeComponent")
	            for each kvpExchangeComponent in kvpExchangeComponents
	
		            for each exchangeDataItem in kvpExchangeComponent.GuestIntrinsicExchangeItems
		                objXMLDoc.loadXML(exchangeDataItem) 
						WScript.Echo objXMLDoc
						xpath = "/INSTANCE/PROPERTY[@NAME='Name']/VALUE[child:text() = 'FullyQualifiedDomainName']"
						set node = objXMLDoc.selectSingleNode(xpath) 
				
						if Not (node Is Nothing) then
						    xpath = "/INSTANCE/PROPERTY[@NAME='Data']/VALUE/child:text()"
						    set node = objXMLDoc.selectSingleNode(xpath) 
						    if Not (node Is Nothing) Then
						    	vmData = addCategoryData(vmData, "Name", node.Text)
								'line = line & vbTab & "dnsName=" & node.Text
						    end if
				        End if
				                                
		                xpath = "/INSTANCE/PROPERTY[@NAME='Name']/VALUE[child:text() = 'NetworkAddressIPv4']"
		                set node = objXMLDoc.selectSingleNode(xpath) 
		                
		                if Not (node Is Nothing) then
		                    xpath = "/INSTANCE/PROPERTY[@NAME='Data']/VALUE/child:text()"
		                    set node = objXMLDoc.selectSingleNode(xpath) 
		                    if Not (node Is Nothing) Then
		                    	vmData = addCategoryData(vmData, "VMIPAddress", node.Text)
								'line = line & vbTab & "ipAddress=" & node.Text
				    		End if
		        		End if
		                
		                xpath = "/INSTANCE/PROPERTY[@NAME='Name']/VALUE[child:text() = 'OSName']"
						set node = objXMLDoc.selectSingleNode(xpath) 
				
						if Not (node Is Nothing) then
						    xpath = "/INSTANCE/PROPERTY[@NAME='Data']/VALUE/child:text()"
						    set node = objXMLDoc.selectSingleNode(xpath) 
						    if Not (node Is Nothing) Then
						    	vmData = addCategoryData(vmData, "OperatingSystem", node.Text)
								'line = line & vbTab & "osName=" & node.Text
						    end if
				        End If
				    Next
				Next
				vmData = vmData&"/>"
			Next
			vmData=vmData&"</VirtualMachines></HyperV>"
			outputText = outputText&vmData
		end if
	end if
end If

'Adding Agent details.
'=====================


agentTaskInfo  = "<AgentTaskInfo><AgentTaskInfo "
agentTaskInfo  = addCategoryData(agentTaskInfo, "AgentTaskID",agentTaskID)
isAgentRunning = isServiceRunning(ae_service)
accountNameReg = ""
siteNameReg = ""
agentId = ""
scanTime = "" &Now
agentSubKey = "Software\ZOHO Corp\ManageEngine AssetExplorer\Agent\"
if(is64BitOS)then
    objReg1.GetStringValue HKEY_LOCAL_MACHINE, agentSubKey, "AgentId", agentId
    objReg1.GetStringValue HKEY_LOCAL_MACHINE, agentSubKey, "Version", version
    objReg1.GetStringValue HKEY_LOCAL_MACHINE, agentSubKey, "AgentPort", port
    objReg1.GetStringValue HKEY_LOCAL_MACHINE, agentSubKey, "AccountName", accountNameReg
    objReg1.GetStringValue HKEY_LOCAL_MACHINE, agentSubKey, "SiteName", siteNameReg
else
objReg.GetStringValue HKEY_LOCAL_MACHINE, agentSubKey, "AgentId", agentId
objReg.GetStringValue HKEY_LOCAL_MACHINE, agentSubKey, "Version", version
objReg.GetStringValue HKEY_LOCAL_MACHINE, agentSubKey, "AgentPort", port
objReg.GetStringValue HKEY_LOCAL_MACHINE, agentSubKey, "AccountName", accountNameReg
objReg.GetStringValue HKEY_LOCAL_MACHINE, agentSubKey, "SiteName", siteNameReg
End if
updateRegistry "LastScanTime", scanTime
agentTaskInfo = addCategoryData(agentTaskInfo,"ComputerName",computerName)
agentTaskInfo = addCategoryData(agentTaskInfo,"ServiceTag",serviceTag)
agentTaskInfo  = addCategoryData(agentTaskInfo, "AgentId",agentId)
agentTaskInfo  = addCategoryData(agentTaskInfo, "MacAddress",macAddress)
agentTaskInfo  = addCategoryData(agentTaskInfo, "LastScanTime",scanTime)
agentTaskInfo  = addCategoryData(agentTaskInfo, "Version",version)
agentTaskInfo  = addCategoryData(agentTaskInfo, "AgentPort",port)
agentTaskInfo  = addCategoryData(agentTaskInfo, "isAgentRunning",isAgentRunning)

agentTaskInfo  = agentTaskInfo & "/></AgentTaskInfo>"
outputText = outputText & agentTaskInfo
If accountNameReg <> "" And siteNameReg <> "" Then
    accountName = accountNameReg
    siteName = siteNameReg
end if

accountInfo = "<Account_Info "
accountInfo = addCategoryData(accountInfo, "AccountName", accountName)
accountInfo = addCategoryData(accountInfo, "SiteName", siteName)
accountInfo = accountInfo & "/>"
outputText = outputText & accountInfo

outputText = outputText & "</DocRoot>"
Err.clear

On Error Resume Next

'Converting Data to XML
'======================
set xml = CreateObject("Microsoft.xmldom")
xml.async = false
loadResult = xml.loadxml(outputText)

On Error Resume Next

hexErrorCode = ""
errordescription = ""
succesMsg = ""
errorMessage = ""
cause = ""
resolution = ""

'Sending Data via http
'=====================
macAddress = Replace(macAddress,"#","")
urlStr = protocol & "://" & hostName & ":" & portNo & "/discoveryServlet/WsDiscoveryServlet?computerName=" & computerName & "&serviceTag=" & serviceTag & "&macAddress=" & macAddress
If (Not isNULL(protocol) and protocol<>"" and InStr(protocol,"https") > 0) then
    set http = createobject("MSXML2.ServerXMLHTTP")
    http.setOption 2, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
Else
    set http = createobject("microsoft.xmlhttp")
End If

if(isAgentMode) Then
    saveXMLFile=true
    computerNameForFile=agentTaskID
elseif (loadResult) Then
        http.open "post",urlStr,false
        http.send xml
        'postErrorMessage()
        if Err Then
            http.send outputText
            if Err Then
                errorMessage = getErrorMessage(Err)
                if(cause = "") then
                    cause = "Could not send the system data to " & protocol & "://" & hostName & ":"&portNo & "."
                end if
                if (not silentMode) Then
                    displayErrorMessage()
                else
                    postErrorMessage()
                End if
                saveXMLFile=true
            else
                succesMsg = "successfully scanned the system data, Find this machine details in AssetExplorer server."
            end if
        else
            succesMsg = "successfully scanned the system data, Find this machine details in AssetExplorer server."
        end if
else 'not agent mode and xml load fails
        http.send outputText
        if Err Then
            errorMessage = getErrorMessage(Err)
            if(cause = "") then
                cause = "Could not send the system data to " & protocol & "://" & hostName & ":"&portNo & "." & newLineConst
            end if
            if (not silentMode) Then
                displayErrorMessage()
            else
                postErrorMessage()
            End if
            saveXMLFile=true
        else
            succesMsg = "successfully scanned the system data, Find this machine details in AssetExplorer server."
        end if
end if


'Saving XML File
'===============

if(saveXMLFile=true) then
    'Saving the Inventory Data as XML File - will be useful to troubleshoot the Error
    fileName = ".\" & computerNameForFile  & ".xml"
    xml.save fileName
    if (Err or getFileSize(fileName) = 0) Then
        writesuccess = writeFile(computerNameForFile & ".xml",outputText)
        if not writesuccess Then
            if(cause = "") then
                errorMessage = getErrorMessage(Err)
                cause = "Could not write the system data in a xml."
                resolution = "Open a command prompt and execute the script as  " & doubleQuote & "cscript ae_scan.vbs -debug >" &computerNameForFile &  ".xml" & doubleQuote &  ",This will generate a file " & doubleQuote & computerNameForFile & ".xml" & doubleQuote &"." & newLineConst & "Send this file to "& supportMailID & " to help you."
                if (not silentMode) Then
                    displayErrorMessage()
                End if

            end if
        end if
    end if
end if
if(isAgentMode) then
    isFile = isFileExist(computerNameForFile &".xml")
    if(isFile And cause = "") then
        succesMsg = "Successfully generated the " & computerNameForFile & ".xml." & "Now You can import the " & computerNameForFile & ".xml in the AssetExplorer server using Standalone audit"
    end if
end if
if (not silentMode) And (succesMsg <> "") then
    WScript.Echo succesMsg
end if

'prints the data for debugmode
'=============================
if (debugmode) Then
    WScript.Echo outputText
end if

'To Add Data
'===========
Function addCategoryData(outputText, category, data)
    'For handling problem when data contains &
    pos=InStr(data,"&")
    if pos>0 Then
        data = Replace(data,"&","###AND###")
    end if
    'For handling problem when data contains <
    pos=InStr(data,"<")
    if pos>0 Then
        data = Replace(data,"<","###[###")
    end if
    'For handling problem when data contains >
    pos=InStr(data,">")
    if pos>0 Then

        data = Replace(data,">","###]###")

    end if
    'For handling problem when data contains DOUBLEQUOTE
    pos=InStr(data,doubleQuote)
    if pos>0 Then
        data = Replace(data,doubleQuote,"###DQ###")
    end if
    data = removeInvalidXMLChar(data)
    retStr = outputText
    if NOT ISNULL(data) then
        retStr = retStr & spaceString
        retStr = retStr & category
        retStr = retStr & equalString
        retStr = retStr & doubleQuote
        retStr = retStr &  Trim(data)
        retStr = retStr & doubleQuote
    end if
    addCategoryData=retStr
End Function

'ConvertToKey method used for fetching windows 8, 8.1 key
Function ConvertToKey(Key)
    Const KeyOffset = 52
    Dim isWin8, Maps, i, j, Current, KeyOutput, Last, keypart1, insert
    isWin8 = (Key(66) \ 6) And 1
    Key(66) = (Key(66) And &HF7) Or ((isWin8 And 2) * 4)
    i = 24
    Maps = "BCDFGHJKMPQRTVWXY2346789"
    Do
       	Current= 0
        j = 14
        Do
           Current = Current* 256
           Current = Key(j + KeyOffset) + Current
           Key(j + KeyOffset) = (Current \ 24)
           Current=Current Mod 24
            j = j -1
        Loop While j >= 0
        i = i -1
        KeyOutput = Mid(Maps,Current+ 1, 1) & KeyOutput
        Last = Current
    Loop While i >= 0 
    keypart1 = Mid(KeyOutput, 2, Last)
    insert = "N"
    KeyOutput = Replace(KeyOutput, keypart1, keypart1 & insert, 2, 1, 0)
    If Last = 0 Then KeyOutput = insert & KeyOutput
    ConvertToKey = Mid(KeyOutput, 1, 5) & "-" & Mid(KeyOutput, 6, 5) & "-" & Mid(KeyOutput, 11, 5) & "-" & Mid(KeyOutput, 16, 5) & "-" & Mid(KeyOutput, 21, 5)
End Function
'To get the licence Key
'======================
Public Function getLicenceKey(bDigitalProductID,version)
    Dim bProductKey()
    Dim bKeyChars(24)
    Dim ilByte
    Dim nCur
    Dim sCDKey
    Dim ilKeyByte
    Dim ilBit
    ReDim Preserve bProductKey(14)
    Set objShell = CreateObject("WScript.Shell")
    Set objShell = Nothing

    if isNull(version)then
        version = 0 ' number less than 14, so that it detects the key for lower versions of office and OS
    end if
    if (version<14) then
        For ilByte = 52 To 66
            bProductKey(ilByte - 52) = bDigitalProductID(ilByte)
        Next
    else
        i=0
        For ilByte = CLng("&h"&328) To CLng("&h"&328)+14
            bProductKey(i) = bDigitalProductID(ilByte)
            i=i+1
        Next
    end if

    bKeyChars(0) = Asc("B")
    bKeyChars(1) = Asc("C")
    bKeyChars(2) = Asc("D")
    bKeyChars(3) = Asc("F")
    bKeyChars(4) = Asc("G")
    bKeyChars(5) = Asc("H")
    bKeyChars(6) = Asc("J")
    bKeyChars(7) = Asc("K")
    bKeyChars(8) = Asc("M")
    bKeyChars(9) = Asc("P")
    bKeyChars(10) = Asc("Q")
    bKeyChars(11) = Asc("R")
    bKeyChars(12) = Asc("T")
    bKeyChars(13) = Asc("V")
    bKeyChars(14) = Asc("W")
    bKeyChars(15) = Asc("X")
    bKeyChars(16) = Asc("Y")
    bKeyChars(17) = Asc("2")
    bKeyChars(18) = Asc("3")
    bKeyChars(19) = Asc("4")
    bKeyChars(20) = Asc("6")
    bKeyChars(21) = Asc("7")
    bKeyChars(22) = Asc("8")
    bKeyChars(23) = Asc("9")
    For ilByte = 24 To 0 Step -1
      nCur = 0
      For ilKeyByte = 14 To 0 Step -1
        nCur = nCur * 256 Xor bProductKey(ilKeyByte)
        bProductKey(ilKeyByte) = Int(nCur / 24)
        nCur = nCur Mod 24
      Next
      sCDKey = Chr(bKeyChars(nCur)) & sCDKey
      If ilByte Mod 5 = 0 And ilByte <> 0 Then sCDKey = "-" & sCDKey
    Next
    getLicenceKey = sCDKey
End Function


'To get Software usage
'=====================
Function getSoftwareUsage(softwareUsageData)
    getSoftwareUsage = "Not Known"
    if not ISNULL(softwareUsageData) then
        usageLevel = CLng(softwareUsageData(24))
        if(usageLevel<3) then
            getSoftwareUsage = "Rarely"
        elseif (usageLevel<9) then
            getSoftwareUsage = "Occasionally"
        elseif (usageLevel<>255) then
            getSoftwareUsage = "Frequently"
        end if
    end if
End Function


'To get the Logical Disk Type
'============================
Function getDiskType(diskType)
    getDiskType="Unknown"
    if(diskType="1") then
        getDiskType="No Root Directory"
    elseif (diskType="2") then
        getDiskType="Removable Disk"
    elseif (diskType="3") then
        getDiskType="Local Disk"
    elseif (diskType="4") then
        getDiskType="Network Drive"
    elseif (diskType="5") then
        getDiskType="Compact Disc"
    elseif (diskType="6") then
        getDiskType="RAM Disk"
    end if
End Function


'To Remove the Index in Network Caption
'======================================
Function getNetworkCaption(captionString)
    getNetworkCaption = captionString
    idx = InStr(captionString," ")
    If(idx>0) Then
        getNetworkCaption = Trim(Mid(captionString,idx))
    End If
End Function

'To Get Monitor Serial number
'============================

Function GetMonitorSerialNumber(EDID)

    sernumstr=""
    sernum=0
    for i=0 to ubound(EDID)-4
        if EDID(i)=0 AND EDID(i+1)=0 AND EDID(i+2)=0 AND EDID(i+3)=255 AND EDID(i+4)=0 Then
            ' if sernum<>0 then
                'sMsgString = "a second serial number has been found!"
                'WScript.ECho sMsgString
                'suspicious=1
            'end if
            sernum=i+4
        end if
    next
    if sernum<>0 then
        endstr=0
        sernumstr=""
        for i=1 to 13
            if EDID(sernum+i)=10 then
                endstr=1
            end if
            if endstr=0 then
                sernumstr=sernumstr & chr(EDID(sernum+i))
            end if

        next
        'sMsgString = "Monitor serial number: " & sernumstr
        'WScript.Echo sMsgString
    else
    sernumstr="-"
    'sMsgString = "No monitor serial number found. Possibly the computer is a laptop."
    'WScript.Echo sMsgString
    end if
    GetMonitorSerialNumber = sernumstr

End Function

'To Handle Error
'===============
Function displayErrorMessage()
    if resolution = "" Then
        resolution = "Open a command prompt and execute the script as  " & doubleQuote & "cscript ae_scan.vbs -debug >" &computerNameForFile &  ".xml" & doubleQuote &  ",This will generate a file " & doubleQuote & computerNameForFile & ".xml" & doubleQuote &"." & newLineConst & "Send this file to "& supportMailID & " to help you."
    end if
    if (not silentMode) Then
        Wscript.Echo errorMessage & newLineConst & "Cause      : " & cause & newLineConst & "Resolution : "& resolution & newLineConst & "If you have any difficulties " & "please report the above Error Message to " & supportMailID
    end if
End Function


'To Get the Error Message for Given Error Code
'=============================================
Function getErrorMessage(Err)
    hexErrorCode = "0x" & hex(Err.Number)
    errordescription = Err.Description
    errorMessage = newLineConst & newLineConst
    errorMessage = errorMessage & "Exception occured while running the Script. (ManageEngine AssetExplorer)"
    errorMessage = errorMessage & newLineConst
    errorMessage = errorMessage & newLineConst & newLineConst

    if(hexErrorCode="0x800C0005") Then
        cause = "The AssetExplorer server is not reachable from this machine."
        resolution = "Check the server name and port number in the script."
    elseif(hexErrorCode="0x80004005") Then
        cause = "The AssetExplorer server is not reachable from this machine."
        resolution = "Check the server name and port number in the script."
    elseif(hexErrorCode="0x80070005") Then
        cause = "The AssetExplorer server is not reachable from this machine."
        resolution = "Check the server name and port number in the script."
    else
        errorMessage = errorMessage & "Error Code : 0x" & hex(Err.Number)
        errorMessage = errorMessage & newLineConst
        errorMessage = errorMessage & "Error Desc : " & Err.description
        errorMessage = errorMessage & newLineConst

    end if
    Err.clear
    errorMessage = errorMessage & newLineConst
    getErrorMessage = errorMessage
End Function


'To post the error message to the server
'=======================================
Function postErrorMessage()
    On Error Resume Next
    if(cause = "") Then
        cause = "-"
    End if
    if(resolution = "") Then
        resolution = "-"
    End if

    exceptionMessage = xmlInfoString &  newLineConst
    exceptionMessage = exceptionMessage &  "<DocRoot>"
    exceptionMessage = exceptionMessage & scriptVersionInfo
    exceptionMessage = exceptionMessage & "<Exception_Msg>"
    exceptionMessage = exceptionMessage & "<Computer "
    exceptionMessage = addCategoryData(exceptionMessage,"Name",computerNameForFile)
    exceptionMessage = exceptionMessage & "/>"
    exceptionMessage = exceptionMessage & "<Error "
    exceptionMessage = addCategoryData(exceptionMessage,"code",hexErrorCode)
    exceptionMessage = addCategoryData(exceptionMessage,"description","errorde scription")
    exceptionMessage = addCategoryData(exceptionMessage,"cause",cause)
    exceptionMessage = addCategoryData(exceptionMessage,"solution",resolution)
    exceptionMessage = exceptionMessage & "/>"
    exceptionMessage = exceptionMessage & "</Exception_Msg>"
    exceptionMessage = exceptionMessage & "<Computer_Info><Computer "
    exceptionMessage = addCategoryData(exceptionMessage,"Name",computerNameForFile)
    exceptionMessage = exceptionMessage & "/></Computer_Info>"
    exceptionMessage = exceptionMessage  & "</DocRoot>"
    'post error message,if it fails write the same in a log file
    http.send exceptionMessage
    if (Err) Then
        sd = writeFile("Error_Scan.log",exceptionMessage)
    End if

End Function

'To write the content in the given filename
'==========================================
Function writeFile(fileName,content)
    On Error Resume Next
    writeFile = false
    Dim oFilesys, oFiletxt, sFilename, sPath
    Set oFilesys = CreateObject("Scripting.FileSystemObject")
    Set oFiletxt = oFilesys.CreateTextFile(fileName,True)
    sPath = oFilesys.GetAbsolutePathName(fileName)
    sFilename = oFilesys.GetFileName(sPath)
    isXPOrLaterOS = isXPAndAbove()
    if(Not isXPOrLaterOS)then
        oFiletxt.WriteLine(content)
        if(Err) then
            writeFile = false
        else
            writeFile = true
        End if
    End if
    oFiletxt.Close'

    if(isXPOrLaterOS)then
        res = saveAsUTF8File(fileName,content)
        if(res) then
            writeFile = false
        else
            writeFile = true
        End if
    End if
End Function

Function isFileExist(fileName)
    On Error Resume Next
    isFileExist = false
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    if (fso.FileExists(fileName)) then
        isFileExist = true
    else
        isFileExist = false
    end if
    set fso = Nothing
End Function

Function getFileSize(fileName)
    On Error Resume Next
    fileSize = 0
    if(isFileExist(fileName))then
	    Dim fso
	    Dim objFile
	    Set fso = CreateObject("Scripting.FileSystemObject")
	    Set objFile = fso.GetFile(fileName)
	    fileSize = objFile.size
	    set objFile = Nothing
	    set fso = Nothing
    end if
    getFileSize = filesize
End Function

'Ref : http://www.w3.org/TR/2000/REC-xml-20001006#NT-Char

Function removeInvalidXMLChar(xmldata)

    Dim isValidChar
    Dim current
    Dim retdata

    retdata = xmldata
    strLen = len(xmldata)
    if(strLen>0)then
        for i=1 to strLen
            current = AscW(Mid(xmldata,i,1))
			IF(current < 0) THEN
				current = 65536 + current
			END IF
            isValidChar = false
            isValidChar = isValidChar or CBool(current = HexToDec("9"))
            isValidChar = isValidChar or CBool(current = HexToDec("A"))
            isValidChar = isValidChar or CBool(current = HexToDec("D"))
            isValidChar = isValidChar or (CBool(current >= HexToDec("20")) and CBool(current <= HexToDec("D7FF")))
            isValidChar = isValidChar or (CBool(current >= HexToDec("E000")) and CBool(current <= HexToDec("FFFD")))
            isValidChar = isValidChar or (CBool(current >= HexToDec("10000")) and CBool(current <= HexToDec("10FFFF")))
            if(Not isValidChar) then
                retdata = Replace(retdata,chr(current),"")
            End if
        Next
    End if
    removeInvalidXMLChar = retdata
End Function

'Hex to decimal
Function HexToDec(hexVal)

    dim dec
    dim strLen
    dim digit
    dim intValue
    dim i

    dec = 0
    strLen = len(hexVal)
    for i =  strLen to 1 step -1

        digit = instr("0123456789ABCDEF", ucase(mid(hexVal, i, 1)))-1
        if digit >= 0 then
                intValue = digit * (16 ^ (len(hexVal)-i))
            dec = dec + intValue
        else
            dec = 0
                i = 0 	'exit for
        end if
    next

  HexToDec = dec
End Function

Function stringTokenizer(strToParse,token)
    Dim res()
    resCount = 0
    if not isNULL(strToParse) and strToParse <> "" then
        do
            n=InStr(strToParse,",")
            if(n>0)then
                resCount = resCount+1
                ReDim Preserve res(resCount)
                res(resCount-1) = Mid(strToParse,1,n-1)

                strToParse = Mid(strToParse,n+1)
            else
                resCount = resCount+1
                ReDim Preserve res(resCount)
                res(resCount-1) = strToParse

            End if
            'n=InStr(str,",")

        Loop while n>0
    End if

    stringTokenizer = res

End  Function

Sub correctUsage
    WScript.Echo  VBCrLf & "USAGE : CSCRIPT ae_scan.vbs [OPTION] " & VBCrLf & VBCrLf & " -SilentMode                                        To run the script in silent mode." & VBCrLf & " -out 'filename'                         To create filename.xml as output" & VBCrLf & " -fs 'file extensions with comma seperated'       To add the files with the given file extensions as a software." & VBCrLf & VBCrLf & "Example: " & "CSCRIPT ae_scan.vbs -fs exe,msi -SilentMode -out scan_data" &VBCrLf
End Sub

Function isServiceRunning(serviceName)

    isServiceRunning = false
    Set objWbemLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objService = objWbemLocator.ConnectServer("localhost", "root\CIMV2")
    Set colItems = objService.ExecQuery("select * from Win32_Service where Name='"&serviceName&"' or DisplayName ='"&serviceName&"'")
    For Each objItem in colItems
        state = objItem.state
        if state = "Running" then
            isServiceRunning = true
        End if
    Next
End Function

Sub updateRegistry(regValue,data)
    if(is64BitOS)then
        objReg1.SetStringValue HKEY_LOCAL_MACHINE, agentSubKey, regValue,data
    else
        objReg.SetStringValue HKEY_LOCAL_MACHINE, agentSubKey, regValue,data
    end if
End Sub

Function saveAsUTF8File( fileName,content)
    On Error Resume Next

    saveAsUTF8File = False
    Dim objStream

    Const adTypeText            = 2
    Const adSaveCreateOverWrite = 2

    if(isXPAndAbove)then   ' ADODB.Stream is applicable for xp and later version only
        Set objStream = CreateObject( "ADODB.Stream" )
        objStream.Open
        objStream.Type = adTypeText
        objStream.Position = 0
        objStream.Charset = xmlEncoding
        objStream.WriteText content
        objStream.SaveToFile fileName, adSaveCreateOverWrite
        objStream.Close
        Set objStream = Nothing

        If Err Then
            saveAsUTF8File = False
        Else
            saveAsUTF8File = True
        End If
    else
        saveAsUTF8File = False
    end if
End Function

Function isXPAndAbove()
    On Error Resume Next
    isXPAndAbove = false
    if(Not isNULL(osVersion) and osVersion<>"")then
        ver = Left(osVersion,3)*1

        if(ver>=5.1)then
            isXPAndAbove = true
        else
            isXPAndAbove = false
        end if
    End if

End Function

Function isVistaAndAbove()
    On Error Resume Next
    isVistaAndAbove = false
    if(Not isNULL(osVersion) and osVersion<>"")then
        ver = Left(osVersion,3)*1

        if(ver>=6.0)then
            isVistaAndAbove = true
        else
            isVistaAndAbove = false
        end if
    End if
End Function

Function isVista()
    On Error Resume Next
    isVista = false
    if(Not isNULL(osVersion) and osVersion<>"")then
        ver = Left(osVersion,3)*1
        if(ver=6.0)then
            isVista = true
        else
            isVista = false
        end if
    End if
End Function

'printers from registry
'======================
Function getPrinterInfo(data,printercount)
    On Error Resume Next
    objReg.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\",providers 'LanMan Print Services\Servers\",arrSubKeys  ' winsys\Printers\Gemini"
    If NOT ISNULL(providers) then
        For Each provider In providers
            objReg.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\"&provider&"\Servers\",servers
            If NOT ISNULL(servers) then
                For Each server In servers
                    objReg.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\"&provider&"\Servers\"&server&"\Printers",printers
                    If NOT ISNULL(printers) then
                        For Each printer In printers
                            objReg.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\"&provider&"\Servers\"&server&"\Printers\"&printer&"\DsSpooler", "PrinterName", printerName
                            if(Not isNULL(printerName) and printerName<>"" and (Not printersList.Exists(printerName))) then
                                count=count+1
                                printersList.add printerName,count
                                data = data & "<Printer_" & count & " "
                                data = addCategoryData(data, "Name", printerName)
                                objReg.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\"&provider&"\Servers\"&server&"\Printers\"&printer&"\DsSpooler", "driverName", printDriver
                                data = addCategoryData(data, "Model", printDriver)
                                objReg.GetMultiStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\"&provider&"\Servers\"&server&"\Printers\"&printer&"\DsSpooler", "portName", printerPorts

                                For Each portName In printerPorts
                                    printerPort = portName
                                Next

                                data = addCategoryData(data, "PortName", printerPort)
                                objReg.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\"&provider&"\Servers\"&server&"\Printers\"&printer&"\DsSpooler", "Location", location
                                data = addCategoryData(data, "Location", location)
                                objReg.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\"&provider&"\Servers\"&server&"\Printers\"&printer&"\DsSpooler", "Description", description
                                data = addCategoryData(data, "Comment", description)
                                objReg.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\"&provider&"\Servers\"&server&"\Printers\"&printer&"\DsSpooler", "serverName", serverName
                                data = addCategoryData(data, "ServerName", serverName)
                                data = addCategoryData(data, "Network", "TRUE")
                                data = addCategoryData(data, "Local", "FALSE")
                                data = data & "/>"
                            End if
                        Next
                    End if

                Next
            End if
        Next
    End if
    getPrinterInfo = data
    Err.clear
End Function

Function updateSQLEdition(softName,regMainKey,regSubKey,regObj)
    if (isSQL(softName))then
        regObj.GetStringValue regMainKey, regSubKey&"\Microsoft SQL Server\MSSQL.1\Setup", "Edition", sqlEdition  'sql 2005 default instance
        if (Not isNull(sqlEdition) And sqlEdition <> "") then
            'WScript.Echo "2005"
            softName = softName&" ("& sqlEdition &")"
        else
            regObj.GetStringValue regMainKey, regSubKey&"\MSSQLServer\Setup", "Edition", sqlEdition	'sql 2000 default instance
            if (Not isNull(sqlEdition) And sqlEdition <> "") then
                'WScript.Echo "2000"
                softName = softName&" ("&sqlEdition&")"
            else
                regObj.GetMultiStringValue regMainKey, regSubKey&"\Microsoft SQL Server\", "InstalledInstances", sqlInsatlledInstances
                'WScript.Echo "3 "&sqlInsatlledInstances(0)
                For Each sqlInsatlledInstance In sqlInsatlledInstances
                    if ( Not isNull(sqlInsatlledInstance) or sqlInsatlledInstance <> "") then
                        regObj.GetStringValue regMainKey, "SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL", sqlInsatlledInstance, sqlInstance
                        'WScript.Echo "4 "&sqlInstance

                        if (Not isNull(sqlInstance) And sqlInstance <> "") then
                            regObj.GetStringValue regMainKey, regSubKey&"\Microsoft SQL Server\"&sqlInstance&"\Setup", "Edition", sqlEdition
                            'WScript.Echo "5 "&sqlEdition
                            if (Not isNull(sqlEdition) And sqlEdition <> "") then
                                softName = softName&" ("&sqlEdition&")"
                                exit for
                            end if
                        end if
                    end if
                Next
            end if
        End if
    End if
    updateSQLEdition = softName
End Function

Function getAdobeLicenseKey(wmiObj,regKey,softName)
    On Error Resume Next
    adobeLicKey = ""

    wmiObj.GetStringValue HKEY_LOCAL_MACHINE,regKey,"EPIC_SERIAL",adobeLicKey
    if( isNULL(adobeLicKey) or Trim(adobeLicKey)="")then
        wmiObj.GetStringValue HKEY_LOCAL_MACHINE,regKey,"SERIAL",adobeLicKey
        if(isNULL(adobeLicKey) or Trim(adobeLicKey)="")then
            verIndex=InstrRev(Trim(softName)," ")
            softNameWithOutVersion = Mid(softName,1,verIndex-1)
            version=Mid(softName,verIndex+1,1)
            wmiObj.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\Adobe\"&softNameWithOutVersion,adobeArrSubKeys1

            If NOT ISNULL(adobeArrSubKeys1) then

                For Each abobeSubkey1 In adobeArrSubKeys1
                    if(inStr(abobeSubkey1,version)=1)then 'if abobeSubkey1 starts with version
                        wmiObj.GetStringValue HKEY_LOCAL_MACHINE,"SOFTWARE\Adobe\"&softNameWithOutVersion&"\"&abobeSubkey1&"\Registration","SERIAL",adobeLicKey
                        if(not ISNULL(adobeLicKey) and adobeLicKey<>"")then
                            exit for
                        else
                            'WScript.Echo "Adobe LicenceKey is not found"
                        End if
                    End if
                Next
            End if
            if(ISNULL(adobeLicKey) or Trim(adobeLicKey)="")then
                productIndex=InStr(Trim(softName), " ")
                productName = Mid(softName,1,productIndex-1)
                softNameWithoutProduct = Mid(softName,productIndex+1,verIndex-productIndex-1)
                wmiObj.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\Adobe\"&softNameWithoutProduct,adobeArrSubKeys2

                If NOT ISNULL(adobeArrSubKeys2) then

                    For Each abobeSubkey2 In adobeArrSubKeys2
                        if(inStr(abobeSubkey2,version)=1)then 'if abobeSubkey2 starts with version

                            wmiObj.GetStringValue HKEY_LOCAL_MACHINE,"SOFTWARE\Adobe\"&softNameWithoutProduct&"\"&abobeSubkey2&"\Registration","SERIAL",adobeLicKey
                            if(not ISNULL(adobeLicKey) and adobeLicKey<>"")then
                                exit for
                            else
                                'WScript.Echo "Adobe LicenceKey is not found"
                            End if
                        End if
                    Next
                End if
            End if
        End if
    End if
    if(ISNULL(adobeLicKey))then
        adobeLicKey=""
    End if
    getAdobeLicenseKey = adobeLicKey
End Function
sub setAutodeskLicenses(objRegLoc)
autodeskKey = "SOFTWARE\Autodesk"
objRegLoc.EnumKey HKEY_LOCAL_MACHINE, autodeskKey, autodeskProducts
If NOT ISNULL(autodeskProducts) then
	For Each adPrd In autodeskProducts
		objRegLoc.EnumKey HKEY_LOCAL_MACHINE, autodeskKey&"\"&adPrd,productdVersions
		if(not ISNULL(productdVersions)) then
			For Each adPrdVer In productdVersions
				objRegLoc.GetStringValue HKEY_LOCAL_MACHINE, autodeskKey&"\"&adPrd&"\"&adPrdVer,"ProductName",adPrdName
				if((ISNULL(adPrdName)) or (adPrdName = ""))then
					objRegLoc.GetStringValue HKEY_LOCAL_MACHINE, autodeskKey&"\"&adPrd&"\"&adPrdVer, "Product Name",adPrdName
				end if
				if((not ISNULL(adPrdName)) and (adPrdName <> ""))then
					objRegLoc.GetStringValue HKEY_LOCAL_MACHINE, autodeskKey&"\"&adPrd&"\"&adPrdVer,"SerialNumber",adSerialNo
					if((not ISNULL(adSerialNo)) and (adSerialNo <> "") and (adSerialNo <> "-") and (instr(adSerialNo, "000-") = 0))then
						autodeskSoftLicKeys.add adPrdName,adSerialNo
						autodeskProductSuiteSerialNumber=adSerialNo
						exit for
					else
						objRegLoc.GetStringValue HKEY_LOCAL_MACHINE, autodeskKey&"\"&adPrd&"\"&adPrdVer,"Serial Number",adSerialNo
						if((not ISNULL(adSerialNo)) and (adSerialNo <> "") and (adSerialNo <> "-") and (instr(adSerialNo, "000-") = 0))then
							autodeskSoftLicKeys.add adPrdName,adSerialNo
							autodeskProductSuiteSerialNumber = adSerialNo
							exit for
						end if
					end if
				end if
				objRegLoc.EnumKey HKEY_LOCAL_MACHINE, autodeskKey&"\"&adPrd&"\"&adPrdVer,adRegKeys
				if (not ISNULL(adRegKeys)) then
					For Each adRegKey In adRegKeys
						objRegLoc.GetStringValue HKEY_LOCAL_MACHINE, autodeskKey&"\"&adPrd&"\"&adPrdVer&"\"&adRegKey,"ProductName",adPrdName
						if((ISNULL(adPrdName)) or (adPrdName = ""))then
							objRegLoc.GetStringValue HKEY_LOCAL_MACHINE, autodeskKey&"\"&adPrd&"\"&adPrdVer&"\"&adRegKey,"Product Name",adPrdName
						end if
						if((not ISNULL(adPrdName)) and (adPrdName <> ""))then
							objRegLoc.GetStringValue HKEY_LOCAL_MACHINE, autodeskKey&"\"&adPrd&"\"&adPrdVer&"\"&adRegKey,"SerialNumber",adSerialNo
							if((not ISNULL(adSerialNo)) and (adSerialNo <> "") and (adSerialNo <> "-") and (instr(adSerialNo, "000-") = 0))then
								autodeskSoftLicKeys.add adPrdName,adSerialNo
								autodeskProductSuiteSerialNumber = adSerialNo
								exit for
							end if
						else
							objRegLoc.GetStringValue HKEY_LOCAL_MACHINE, autodeskKey&"\"&adPrd&"\"&adPrdVer&"\"&adRegKey,"Serial Number",adSerialNo

							if((not ISNULL(adSerialNo)) and (adSerialNo <> "") and (adSerialNo <> "-") and (instr(adSerialNo, "000-") = 0))then
								autodeskSoftLicKeys.add adPrdName,adSerialNo
								autodeskProductSuiteSerialNumber = adSerialNo
								exit for
							end if
						end if
					Next
				end if
			Next
		end if
	Next
end if
end sub

function isSQL(softwareName)
	if (sqlSoftList.Exists(Lcase(softwareName)))then
		isSQL = true
	else
		isSQL = false
	end if
end function

function GetInstallerKeyNameFromGuid(subKeyName)
	installerKeyName = Replace(subKeyName,"{","")
	installerKeyName = Replace(installerKeyName,"}","")
	installerKeyNameList=Split(installerKeyName,"-")
	resultKey=""
	tempCount=0
		for each tempinstallerKeyName in installerKeyNameList
				if tempCount < 3 then
					resultKey = resultKey & "" & StrReverse(tempinstallerKeyName)				
				else
					x=Len(tempinstallerKeyName)
					For i=0 to x-1 
						resultKey= resultKey & Mid(tempinstallerKeyName,i+2,1)
						resultKey= resultKey & Mid(tempinstallerKeyName,i+1,1)
						i = i+1
					Next   
				end if
				tempCount = tempCount + 1
		next
		GetInstallerKeyNameFromGuid = resultKey
end function
sub initSQLSoftList
	sqlSoftList.add "microsoft sql server 2000", ""
	sqlSoftList.add "microsoft sql server 2005", "" 
	sqlSoftList.add "microsoft sql server 2008", "" 
	sqlSoftList.add "microsoft sql server 2008 r2", "" 
	sqlSoftList.add "microsoft sql server 2012", "" 
	sqlSoftList.add "microsoft sql server 2000 (64-bit)", "" 
	sqlSoftList.add "microsoft sql server 2005 (64-bit)", "" 
	sqlSoftList.add "microsoft sql server 2008 (64-bit)", "" 
	sqlSoftList.add "microsoft sql server 2008 r2 (64-bit)", "" 
	sqlSoftList.add "microsoft sql server 2012 (64-bit)", "" 
end sub
Sub printLog(msg)
	'WScript.Echo msg
End Sub
Function readValueFromEDID(matchingArray,indexArray,rawEDIDArray)
  Dim idx, matchFound, strTemp
  For i=0 To UBOUND(indexArray)
	  matchFound  = True
	  For idx = 0 To 3
		If CInt( matchingArray( idx )  ) <> CInt( rawEDIDArray( idx + indexArray(i) ) ) Then matchFound  = False
	  Next
	  If matchFound Then
		For idx = 4 To 17
		  Select Case rawEDIDArray( indexArray(i) + idx )
			Case 0
			  strTemp = strTemp & " "
			Case 7
			  strTemp = strTemp & " "
			Case 10
			  strTemp = strTemp & " "
			Case 13
			  strTemp = strTemp & " "
			Case Else
			  strTemp = strTemp & Chr( rawEDIDArray( indexArray(i) + idx ) )
		  End Select
		Next
		strTemp = Trim( strTemp )
		readValueFromEDID = strTemp
	  End If
	Next
End Function

Function isAnalog(rawEDIDArray)
    bitValue = Mid(DecimalToBinary(rawEDIDArray(20)),1,1)
    tempVar = ""
    If (bitValue = "1") Then
        tempVar = "Digital"
    Else
        tempVar = "Analog"
    End If
    isAnalog = tempVar
End Function
Function getMonitorSize(rawEDIDArray)
    hor_resolution = rawEDIDArray(54 + 2) + (rawEDIDArray(54 + 4) And 240) * 16 
    vert_resolution = rawEDIDArray(54 + 5) + (rawEDIDArray(54 + 7) And 240) * 16
    width = rawEDIDArray(21)
    height = rawEDIDArray(22)
    diagonal = Sqr(((width)*(width))+((height)*(height)))
    getMonitorSize = Round((diagonal * 0.39370),1) & """ (" & hor_resolution & "x" & vert_resolution & ")"
End Function
Function getProductCode(rawEDIDArray)
    strTemp =""
    strTemp = Hex(rawEDIDArray (11)) & Hex(rawEDIDArray (10))
    getProductCode = strTemp
End Function
Function getManufacturerCode(rawEDIDArray)
    binaryString = DecimalToBinary(rawEDIDArray(8)) & DecimalToBinary(rawEDIDArray(9))
    mfgCode = ""
    tempStr = ""
    count = 0
    For i=2 To Len(binaryString)
        tempStr = tempStr & Mid(binaryString,i,1)
        count = count + 1
        If count = 5 Then
            mfgCode = mfgCode & Chr(64 + BinaryToDecimal(tempStr))
            tempStr = ""
            count = 0
        End If
    Next
    getManufacturerCode = mfgCode
End Function

Function BinaryToDecimal(Binary)  
  For s = 1 To Len(Binary)
    n = n + (Mid(Binary, Len(Binary) - s + 1, 1) * (2 ^ (s - 1)))
  Next  
  BinaryToDecimal = n
End Function
Function DecimalToBinary(DecimalNum)
    count=1
    n = CLng(DecimalNum)
    tmp = (n Mod 2)
    n = n \ 2       
    Do While n <> 0
    tmp = (n Mod 2) & tmp
    count = count+1
    n = n \ 2
    Loop
    For i=count to 7
        tmp = "0" & tmp
    Next
    DecimalToBinary = tmp
End Function

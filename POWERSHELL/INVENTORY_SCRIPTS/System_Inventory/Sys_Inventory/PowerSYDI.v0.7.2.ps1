<#
.SYNOPSIS
Extracts information about the operating system, applications installed and 
hardware for a windows system and creating a html or XML file.

.DESCRIPTION
Extracts information about the Windows server or workstation, for example 
operating system, applications installed and hardware etc and generates a report
in html file or an XML file.  The files generated will have the 
same name as the server in question and the corresponding extension for 
the relevant output type.  

.PARAMETER Output 
What format should the output be generated in, either html or XML. 
Default: html

.PARAMETER Computername
The name of the computer you wish to run against, IP addresses can also be
user. Defaults to the local computer.

.PARAMETER LoadDoc
Display the output once the script has completed

.PARAMETER NoPing
Do not ping servers first, useful for servers that have the Windows Firewall 
enabled. Default: False

.PARAMETER path
Path to where you wish to save the output. Default: ".\", current folder

.PARAMETER username 
Username the script should run under, such as “User01”, “Domain01\User01”, or 
User@domain.com.   If you have not used –password, you will be prompted for a 
password.

.PARAMETER password 
Used in conjunction with the above –username

.PARAMETER Credential
Powershell credential object you wish to run the script under, cannot be used 
with above username and password

.PARAMETER reportSoftware, reportHardware, reportStorage, reportNetwork, 
               reportMisc, reportExtra, ReportUser
Extract corresponding details and add to report. Default: True

.EXAMPLE 

PowerSYDI.ps1
Reports on current system and save to current folder, using your logged on credentials.

.EXAMPLE 
PowerSYDI.ps1 -computer server2 –username domain\admininstrator -noping:$true
Runs the script against host server2, using the domain administrator account. You will be prompted for a password and the server will not be pinged. The output will be called server2.html

.EXAMPLE 
PowerSYDI.ps1 -computer server3 -path d:\documents -LoadDoc:$true
Runs the script against host server3, saving the output to d:\documents and display the html webpage. The output will be called server3.html

.EXAMPLE 
PowerSYDI.ps1 -computer server3 -LoadDoc:$true -reportMisc:$false -reportExtra:$false
Runs the script against host server4, display the output and don't report on miscellaneous or extra information. The output will be called server4.html

.LINK
http://www.powerforge.net

#>
param (
	[string]$Output = "html",
	[string]$Computer = ".",
	[Boolean]$LoadDoc = $False,
	[string]$Username = "",  
	[string]$Password = "",  
	[string]$path = ".\",
	[Boolean]$NoPing = $True,
	[Object]$Credential = $null,
	[Boolean]$ReportSoftware = $True,  
	[Boolean]$ReportHardware = $True,
	[Boolean]$ReportStorage = $True,
	[Boolean]$ReportNetwork = $True,
	[Boolean]$ReportUser = $True,
	[Boolean]$ReportExtra = $True,
	[Boolean]$ReportMisc = $True
)

# Copyright (c) 2004-2009 Patrick Ogenstad
# Copyright (c) 2013-2016 Carl Armstrong
# All rights reserved.
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
#  * Redistributions of source code must retain the above copyright notice,
#    this list of conditions and the following disclaimer.
#  * Redistributions in binary form must reproduce the above copyright notice,
#    this list of conditions and the following disclaimer in the documentation
#    and/or other materials provided with the distribution.
#  * Neither the name SYDI nor the names of its contributors may be used
#    to endorse or promote products derived from this software without
#    specific prior written permission.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
# AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
# ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE
# LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
# CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
# SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
# INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
# CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
# ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
# POSSIBILITY OF SUCH DAMAGE.

#Version 
$Version = "0.7.2"
# To do
# Tidy up variables/display names
# Arguments...
# IIS
# Add more html sytles

# Test!!!!!!

#region Define variables
	$Win32Hardware = "Win32_Processor", "Win32_BIOS", "Win32_ComputerSystemProduct", "Win32_SystemEnclosure", "Win32_PhysicalMemory", "Win32_CDROMDrive", 
	                          "Win32_SoundDevice", "Win32_VideoController", "Win32_TapeDrive","Win32_BaseBoard","Win32_Keyboard","Win32_PointingDevice"
	$Win32Software = "Win32_OptionalFeature", "Win32_QuickFixEngineering", "Win32_Product"   
	$Win32Storage  = "Win32_DiskDrive", "Win32_DiskDriveToDiskPartition", "Win32_DiskPartition", "Win32_LogicalDisk", "Win32_LogicalDiskToPartition"
	$Win32Network  = "Win32_NetworkAdapterConfiguration", "Win32_IP4RouteTable"
	$Win32User     = "Win32_Group#LocalAccount=True#", "Win32_UserAccount#LocalAccount=True#","Win32_GroupUser"
	$Win32Misc     = "Win32_TimeZone","Win32_TCPIPPrinterPort","Win32_PrinterDriver", "Win32_NTEventLogFile", "Win32_Printer",
	                          "Win32_Process", "Win32_Service", "Win32_Share", "Win32_StartupCommand", "Win32_PageFile", "Win32_Registry",
	                                        "Win32_Environment","Win32_NTDomain"
	$Win32Extra    = "MicrosoftNLB_Cluster##root\MicrosoftNLB", "MicrosoftNLB_ClusterSetting##root\MicrosoftNLB", "MicrosoftNLB_Node##root\MicrosoftNLB",
			                                "MicrosoftNLB_NodeSetting##root\MicrosoftNLB","MicrosoftNLB_PortRuleEx##root\MicrosoftNLB",
					 "MSCluster_Cluster##root\mscluster","MSCluster_Node##root\mscluster","MSCluster_Resource##root\mscluster",
					 "MSCluster_Network##root\mscluster","MSCluster_DiskPartition##root\mscluster","MSCluster_ResourceGroup##root\mscluster",
					 "MSCluster_ResourceToPossibleOwner##root\mscluster","MSCluster_ResourceGroupToPreferredNode##root\mscluster",
					 "MSCluster_ResourceGroupToResource##root\mscluster","MSCluster_ResourceToDisk##root\mscluster",
					 "MSCluster_DiskToDiskPartition##root\mscluster"
	$i = 0
	$vbCrLf = "`r`n"
	$vbTab =  "`t" 
	$Quotes = [char]34
	$HKLM = 2147483650

	$Ping = $null
#EndRegion  Define Constants

#region Style
$style = "<style>
      /*******************************************************
      TITLE: Fluid Two-Column Layout (Basic) V1.0 (Beta)
      DATE: 20060418
      AUTHOR: The CSS Tinderbox - http://www.csstinderbox.com
      *******************************************************/

      body {margin:0;padding:0;background-color:#ffffff;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:1em;color: white;}
      a{text-decoration:none; color:#000000;}
      blockquote {margin:1em;padding:.5em;font-size:.9em;background-color:#F3F2ED;border-top:1px solid #999999;border-bottom:1px solid #999999;}
      blockquote p {margin:.2em;}
      #header {margin:2em 2em 0 2em;padding:1em 1.5em;height:5em;background-color:#A30B1A;border:2px solid #eeeeee;}
      #header h1 { margin:0; padding:0;font-size:1.5em;}
      #header h2 { margin:0; padding:0;font-size:1.2em;}
      #header h3 { margin:0; padding:0;font-size:.9em;}
      #footer {margin:2em 2em 0 2em;padding:1em 1.5em;height:2em;background-color:#A30B1A;border:2px solid #eeeeee;}
      #footer h3 { margin:0; padding:0;font-size:.9em;color:#FFFFFF}
      #leftColumn {position:absolute;left:2.25em;top:10.3em;width:8em;margin:0;padding:1em .5em 2em .5em;background-color:#A30B1A;border:1px solid #eeeeee;font-size:.9em;}
      #leftColumn h2 { margin:0 0 -1em 0; padding:0;font-size:1em;letter-spacing:.1em;}
      #leftColumn ul { margin:1.5em 0 0 0; padding:0;list-style:none;}
      #leftColumn li { margin:0 0 .4em 0; padding:0;}
      #leftColumn li a { margin:0 0 0 .2em;color: white}
      #centerColumn {right:2.25em;margin-top:.2em;margin-left: 12.50em;margin-right:2.25em;voice-family: ""\""}\"""";voice-family: inherit;margin-right:2.25em;padding:1em .5em 2em .5em;background:#FFFFFF;font-size:.9em;}
      #centerColumn h1 { margin:0 0 -1em 0; padding:0;font-size:1.5em;letter-spacing:.1em;}
      #centerColumn h2 { margin:0 0 -1em 0; padding:0;font-size:1.2em;letter-spacing:.1em; padding-top:.5em;padding-bottom:.5em;}
      #centerColumn h3 { margin:0 0 -1em 0; padding:0;font-size:0.9em;letter-spacing:.1em; padding-top:.5em;padding-bottom:.5em;}
      #tags {margin:0 0 .5em 0;width:10em;float:left;border:none;text-align:left;background-color:#A30B1A;}
      #tags img {border:none;}
      #tags p {margin:0 0 .25em 0;}
      #tags a {font-size:.7em;}

      #table1 {font-family:""Trebuchet MS"", Arial, Helvetica, sans-serif;width:100%;border-collapse:collapse;}
      #table1 td, #table1 th {font-size:1em;border:1px solid #F6F0E0;padding:3px 7px 2px 7px;}
      #table1 th {font-size:1.1em;text-align:left;padding-top:5px;padding-bottom:4px;background-color:#CFB59F;color:#ffffff;}
      #table1 td.alt {font-size:1.1em;text-align:left;padding-top:5px;padding-bottom:4px;background-color:#F3F2ED;color:#000000;}
      #table1 th {font-size:1em;border:1px solid #F6F0E0;padding:3px 7px 2px 7px;}
      #table1 tr.alt td {color:#000000;background-color:#F3F2ED;}
</style>"
#endregion Style

#region FUNCTIONS

#region Win32_TCPIPPrinterPort Functions

function get_Win32_TCPIPPrinterPort_Protocol($intProtocol)
{
    switch ($intProtocol)
    {
        1  { $ret = "RAW"}
        2  { $ret = "LPR"}
        default { $ret = $intProtocol }
    }
    return $ret
}

#endregion 
 
#region Win32_Share Functions

function get_Win32_Share_Type($intType)
{
    switch ($intType)
    {
            0  { $ret = "Disk Drive"}
            1  { $ret = "Print Queue"}
            2  { $ret = "Device"}
            3  { $ret = "IPC"}
            2147483648  { $ret = "Disk Drive Admin"}
            2147483649  { $ret = "Print Queue Admin"}
            2147483650  { $ret = "Device Admin"}
            2147483651  { $ret = "IPC Admin"}
            default { $ret = $intType }
      }
      return $ret
}

#endregion 

#region Win32_SystemEnclosure Functions
      
      function Get-Win32_SystemEnclosure-ChassisTypes($uint16_array_ChassisTypes)
      {
            switch ($uint16_array_ChassisTypes)
            {
                  1 { $ret = "Other"}
                  2 { $ret = "Unknown"}
                  3 { $ret = "Desktop"}
                  4 { $ret = "Low Profile Desktop"}
                  5 { $ret = "Pizza Box"}
                  6 { $ret = "Mini Tower"}
                  7 { $ret = "Tower"}
                  8 { $ret = "Portable"}
                  9 { $ret = "Laptop"}
                 10 { $ret = "Notebook"}
                  11 { $ret = "Hand Held"}
                  12 { $ret = "Docking Station"}
                  13 { $ret = "All in One"}
                  14 { $ret = "Sub Notebook"}
                  15 { $ret = "Space-Saving"}
                  16 { $ret = "Lunch Box"}
                  17 { $ret = "Main System Chassis"}
                  18 { $ret = "Expansion Chassis"}
                  19 { $ret = "SubChassis"}
                  20 { $ret = "Bus Expansion Chassis"}
                  21 { $ret = "Peripheral Chassis"}
                  22 { $ret = "Storage Chassis"}
                  23 { $ret = "Rack Mount Chassis"}
                  24 { $ret = "Sealed-Case PC"}
                  default { $ret = $uint16_array_ChassisTypes }
            }
            return $ret
      }

      function Get-Win32_SystemEnclosure-SecurityBreach($uint16_SecurityBreach)
      {
            switch ($uint16_SecurityBreach)
            {
                  1 { $ret = "Other"}
                  2 { $ret = "Unknown"}
                  3 { $ret = "No Breach"}
                  4 { $ret = "Breach Attempted"}
                  5 { $ret = "Breach Successful"}
                  default { $ret = $uint16_SecurityBreach }
            }
            return $ret
      }

      function Get-Win32_SystemEnclosure-SecurityStatus($uint16_SecurityStatus)
      {
            switch ($uint16_SecurityStatus)
            {
                  1 { $ret = "Other"}
                  2 { $ret = "Unknown"}
                  3 { $ret = "None"}
                  4 { $ret = "External Interface Locked Out"}
                  5 { $ret = "External Interface Enabled"}
                  default { $ret = $uint16_SecurityStatus }
            }
            return $ret
      }

      function Get-Win32_SystemEnclosure-ServicePhilosophy($uint16_array_ServicePhilosophy)
      {
            switch ($uint16_array_ServicePhilosophy)
            {
                  0 { $ret = "Unknown"}
                  1 { $ret = "Other"}
                  2 { $ret = "Service From Top"}
                  3 { $ret = "Service From Front"}
                  4 { $ret = "Service From Back"}
                  5 { $ret = "Service From Side"}
                  6 { $ret = "Sliding Trays"}
                  7 { $ret = "Removable Sides"}
                  8 { $ret = "Moveable"}
                  default { $ret = $uint16_array_ServicePhilosophy }
            }
            return $ret
      }

      #endregion Win32_SystemEnclosure Functions

#region Win32_NetworkAdapter Functions

function get_Win32_NetworkAdapter_Availability($intAvailability)
{
    switch ($intAvailability)
    {
            1  { $ret = "Other"}
            2  { $ret = "Unknown"}
            3  { $ret = "Running or Full Power"}
            4  { $ret = "Warning"}
            5  { $ret = "In Test"}
            6  { $ret = "Not Applicable"}
            7  { $ret = "Power Off"}
            8  { $ret = "Off Line"}
            9  { $ret = "Off Duty"}
            10  { $ret = "Degraded"}
            11  { $ret = "Not Installed"}
            12  { $ret = "Install Error"}
            13  { $ret = "Power Save - Unknown"}
            14  { $ret = "Power Save - Low Power Mode"}
            15  { $ret = "Power Save - Standby"}
            16  { $ret = "Power Cycle"}
            17  { $ret = "Power Save - Warning"}

           default { $ret = $intAvailability }
      }
      return $ret
}

function get_Win32_NetworkAdapter_NetConnectionStatus($intNetConnectionStatus)
{
    switch ($intNetConnectionStatus)
    {

           0  { $ret = "Disconnected"}
            1  { $ret = "Connecting"}
            2  { $ret = "Connected"}
            3  { $ret = "Disconnecting"}
            4  { $ret = "Hardware not present"}
            5  { $ret = "Hardware disabled"}
            6  { $ret = "Hardware malfunction"}
            7  { $ret = "Media disconnected"}
            8  { $ret = "Authenticating"}
            9  { $ret = "Authentication succeeded"}
            10  { $ret = "Authentication failed"}
            11  { $ret = "Invalid address"}
            12  { $ret = "Credentials required"}
            default { $ret = $intNetConnectionStatus }
      }
      return $ret
}
#endregion

#region Win32_BIOS Functions
function Get-Win32_BIOS-BiosCharacteristics($uint16_array_BiosCharacteristics)
{
      switch ($uint16_array_BiosCharacteristics)
      {
            0 { $ret = "Reserved"}
            1 { $ret = "Reserved"}
            2 { $ret = "Unknown"}
            3 { $ret = "BIOS Characteristics Not Supported"}
            4 { $ret = "ISA is supported"}
            5 { $ret = "MCA is supported"}
            6 { $ret = "EISA is supported"}
            7 { $ret = "PCI is supported"}
            8 { $ret = "PC Card (PCMCIA) is supported"}
            9 { $ret = "Plug and Play is supported"}
            10 { $ret = "APM is supported"}
            11 { $ret = "BIOS is Upgradable (Flash)"}
            12 { $ret = "BIOS shadowing is allowed"}
            13 { $ret = "VL-VESA is supported"}
            14 { $ret = "ESCD support is available"}
            15 { $ret = "Boot from CD is supported"}
            16 { $ret = "Selectable Boot is supported"}
            17 { $ret = "BIOS ROM is socketed"}
            18 { $ret = "Boot From PC Card (PCMCIA) is supported"}
            19 { $ret = "EDD (Enhanced Disk Drive) Specification is supported"}
            20 { $ret = "Int 13h - Japanese Floppy for NEC 9800 1.2mb (3.5, 1k Bytes/Sector, 360 RPM) is supported"}
            21 { $ret = "Int 13h - Japanese Floppy for Toshiba 1.2mb (3.5, 360 RPM) is supported"}
            22 { $ret = "Int 13h - 5.25 / 360 KB Floppy Services are supported"}
            23 { $ret = "Int 13h - 5.25 /1.2MB Floppy Services are supported"}
            24 { $ret = "Int 13h - 3.5 / 720 KB Floppy Services are supported"}
            25 { $ret = "Int 13h - 3.5 / 2.88 MB Floppy Services are supported"}
            26 { $ret = "Int 5h, Print Screen Service is supported"}
            27 { $ret = "Int 9h, 8042 Keyboard services are supported"}
            28 { $ret = "Int 14h, Serial Services are supported"}
            29 { $ret = "Int 17h, printer services are supported"}
            30 { $ret = "Int 10h, CGA/Mono Video Services are supported"}
            31 { $ret = "NEC PC-98"}
            32 { $ret = "ACPI is supported"}
            33 { $ret = "USB Legacy is supported"}
            34 { $ret = "AGP is supported"}
            35 { $ret = "I2O boot is supported"}
            36 { $ret = "LS-120 boot is supported"}
            37 { $ret = "ATAPI ZIP Drive boot is supported"}
            38 { $ret = "1394 boot is supported"}
            39 { $ret = "Smart Battery is supported"}
            default { $ret = $uint16_array_BiosCharacteristics }
      }
      return $ret
}

function Get-Win32_BIOS-SoftwareElementState($uint16_SoftwareElementState)
{
      switch ($uint16_SoftwareElementState)
      {
            0 { $ret = "Deployable"}
            1 { $ret = "Installable"}
            2 { $ret = "Executable"}
            3 { $ret = "Running"}
            default { $ret = $uint16_SoftwareElementState }
      }
      return $ret
}

function Get-Win32_BIOS-TargetOperatingSystem($uint16_TargetOperatingSystem)
{
      switch ($uint16_TargetOperatingSystem)
      {
            0 { $ret = "Unknown"}
            1 { $ret = "Other"}
            2 { $ret = "MACOS"}
            3 { $ret = "ATTUNIX"}
            4 { $ret = "DGUX"}
            5 { $ret = "DECNT"}
            6 { $ret = "Digital Unix"}
            7 { $ret = "OpenVMS"}
            8 { $ret = "HPUX"}
            9 { $ret = "AIX"}
            10 { $ret = "MVS"}
            11 { $ret = "OS400"}
            12 { $ret = "OS/2"}
            13 { $ret = "JavaVM"}
            14 { $ret = "MSDOS"}
            15 { $ret = "WIN3x"}
            16 { $ret = "WIN95"}
            17 { $ret = "WIN98"}
            18 { $ret = "WINNT"}
            19 { $ret = "WINCE"}
            20 { $ret = "NCR3000"}
            21 { $ret = "NetWare"}
            22 { $ret = "OSF"}
            23 { $ret = "DC/OS"}
            24 { $ret = "Reliant UNIX"}
            25 { $ret = "SCO UnixWare"}
            26 { $ret = "SCO OpenServer"}
            27 { $ret = "Sequent"}
            28 { $ret = "IRIX"}
            29 { $ret = "Solaris"}
            30 { $ret = "SunOS"}
            31 { $ret = "U6000"}
            32 { $ret = "ASERIES"}
            33 { $ret = "TandemNSK"}
            34 { $ret = "TandemNT"}
            35 { $ret = "BS2000"}
            36 { $ret = "LINUX"}
            37 { $ret = "Lynx"}
            38 { $ret = "XENIX"}
            39 { $ret = "VM/ESA"}
            40 { $ret = "Interactive UNIX"}
            41 { $ret = "BSDUNIX"}
            42 { $ret = "FreeBSD"}
            43 { $ret = "NetBSD"}
            44 { $ret = "GNU Hurd"}
            45 { $ret = "OS9"}
            46 { $ret = "MACH Kernel"}
            47 { $ret = "Inferno"}
            48 { $ret = "QNX"}
            49 { $ret = "EPOC"}
            50 { $ret = "IxWorks"}
            51 { $ret = "VxWorks"}
            52 { $ret = "MiNT"}
            53 { $ret = "BeOS"}
            54 { $ret = "HP MPE"}
            55 { $ret = "NextStep"}
            56 { $ret = "PalmPilot"}
            57 { $ret = "Rhapsody"}
            58 { $ret = "Windows 2000"}
            59 { $ret = "Dedicated"}
            60 { $ret = "VSE"}
            61 { $ret = "TPF"}
            default { $ret = $uint16_TargetOperatingSystem }
      }
      return $ret
}

#endregion Win32_BIOS Functions

#region Win32_ComputerSystem Functions

function get_Win32_ComputerSystem_PowerState($intPowerState)
{
    switch ($intPowerState)
    {
            0  { $ret = "Unknown"}
            1  { $ret = "Full Power"}
            2  { $ret = "Power Save - Low Power Mode"}
            3  { $ret = "Power Save - Standby"}
            4  { $ret = "Power Save - Unknown"}
            5  { $ret = "Power Cycle"}
            6  { $ret = "Power Off"}
            7  { $ret = "Power Save - Warning"}

           default { $ret = $intPowerState }
      }
      return $ret
}


function get_Win32_ComputerSystem_ChassisBootupState($intChassisBootupState)
{
    switch ($intChassisBootupState)
    {
            1  { $ret = "Other"}
            2  { $ret = "Unknown"}
            3  { $ret = "Safes"}
            4  { $ret = "Warning"}
            5  { $ret = "Critical"}
            6  { $ret = "Nonrecoverable"}
            default { $ret = $intChassisBootupState }
      }
      return $ret
}

function get_Win32_ComputerSystem_PowerSupplyState($intPowerSupplyState)
{
    switch ($intPowerSupplyState)
    {
            1  { $ret = "Other"}
            2  { $ret = "Unknown"}
            3  { $ret = "Safe"}
            4  { $ret = "Warning"}
            5  { $ret = "Critical"}
            6  { $ret = "Nonrecoverable"}
            default { $ret = $intPowerSupplyState }
      }
      return $ret
}

function get_Win32_ComputerSystem_DomainRole($intDomainRole)
{
    switch ($intDomainRole)
    {
        0 { $ret = "Standalone Workstation"}
        1 { $ret = "Member Workstation"}
        2 { $ret = "Standalone Server"}
        3 { $ret = "Member Server"}
        4 { $ret = "Backup Domain Controller"}
        5 { $ret = "Primary Domain Controller"}
        default { $ret = $intDomainRole }
    }
    return $ret
}

function get_Win32_ComputerSystem_PCSystemType($intPCSystemType)
{
    switch ($intPCSystemType)
    {
        0 { $ret = "Unspecified"}
        1 { $ret = "Desktop"}
        2 { $ret = "Mobile"}
        3 { $ret = "Workstation"}
        4 { $ret = "Enterprise Server"}
        5 { $ret = "Small Office and Home Office (SOHO) Server"}
        6 { $ret = "Appliance PC"}
        7 { $ret = "Performance Server"}
        8 { $ret = "Maximum"}
        default { $ret = $intPCSystemType }
    }
    return $ret
}
#endregion 
 
#region Win32_OperatingSystems


Function Get-Win32_OperatingSystem-ProductType($uint32_ProductType)
{
      switch ($uint32_ProductType)
      {
            1 { $ret = "Work Station"}
            2 { $ret = "Domain Controller"}
            3 { $ret = "Server"}
            default { $ret = $uint32_ProductType }
      }
      return $ret
}

Function Get-Win32_OperatingSystem-SuiteMask($uint32_SuiteMask)
{
      switch ($uint32_SuiteMask)
      {
            1 { $ret = "Small Business"}
            2 { $ret = "Enterprise"}
            4 { $ret = "BackOffice"}
            8 { $ret = "Communications"}
            16 { $ret = "Terminal"}
            32 { $ret = "Small Business Restricted"}
            64 { $ret = "Embedded NT"}
            128 { $ret = "Data Center"}
            256 { $ret = "Single User"}
            512 { $ret = "Personal"}
            1024 { $ret = "Blade"}
            default { $ret = $uint32_SuiteMask }
      }
      return $ret
}

function get_Win32_OperatingSystem_ForegroundApplicationBoost($intForegroundApplicationBoost )
{
    switch ($intForegroundApplicationBoost )
    {
            0 { $ret = "None"}
            1 { $ret = "Minimum"}
            2 { $ret = "Maximum"}
            default { $ret = $intForegroundApplicationBoost }
      }
      return $ret
}

function get_Win32_OperatingSystem_OperatingSystemSKU($intOperatingSystemSKU )
{
    switch ($intOperatingSystemSKU )
    {
            1 { $ret = "Ultimate"}
            2 { $ret = "Home Basic"}
            3 { $ret = "Home Premium"}
            4 { $ret = "Enterprise"}
            5 { $ret = "Home Basic N"}
            6 { $ret = "Business"}
            7 { $ret = "Server Standard"}
            8 { $ret = "Server Datacenter (full installation)"}
            9 { $ret = "Windows Small Business Server"}
            10 { $ret = "Server Enterprise (full installation)"}
            11 { $ret = "Starter"}
            12 { $ret = "Server Datacenter (core installation)"}
            13 { $ret = "Server Standard (core installation)"}
            14 { $ret = "Server Enterprise (core installation)"}
            15 { $ret = "Server Enterprise for Itanium-based Systems"}
            16 { $ret = "Business N"}
            17 { $ret = "Web Server (full installation)"}
            18 { $ret = "HPC Edition"}
            19 { $ret = "Windows Storage Server 2008 R2 Essentials"}
            20 { $ret = "Storage Server Express"}
            21 { $ret = "Storage Server Standard"}
            22 { $ret = "Storage Server Workgroup"}
            23 { $ret = "Storage Server Enterprise"}
            24 { $ret = "Windows Server 2008 for Windows Essential Server Solutions"}
            25 { $ret = "Small Business Server Premium"}
            26 { $ret = "Home Premium N"}
            27 { $ret = "Enterprise N"}
            28 { $ret = "Ultimate N"}
            29 { $ret = "Web Server (core installation)"}
            30 { $ret = "Windows Essential Business Server Management Server"}
            31 { $ret = "Windows Essential Business Server Security Server"}
            32 { $ret = "Windows Essential Business Server Messaging Server"}
            33 { $ret = "Server Foundation"}
            34 { $ret = "Windows Home Server 2011"}
            35 { $ret = "Windows Server 2008 without Hyper-V for Windows Essential Server Solutions"}
            36 { $ret = "Server Standard without Hyper-V"}
            37 { $ret = "Server Datacenter without Hyper-V (full installation)"}
            38 { $ret = "Server Enterprise without Hyper-V (full installation)"}
            39 { $ret = "Server Datacenter without Hyper-V (core installation)"}
            40 { $ret = "Server Standard without Hyper-V (core installation)"}
            41 { $ret = "Server Enterprise without Hyper-V (core installation)"}
            42 { $ret = "Microsoft Hyper-V Server"}
            43 { $ret = "Storage Server Express (core installation)"}
            44 { $ret = "Storage Server Standard (core installation)"}
            45 { $ret = "Storage Server Workgroup (core installation)"}
            46 { $ret = "Storage Server Enterprise (core installation)"}
            47 { $ret = "Starter N"}
            48 { $ret = "Professional"}
            49 { $ret = "Professional N"}
            50 { $ret = "Windows Small Business Server 2011 Essentials"}
            51 { $ret = "Server For SB Solutions"}
            52 { $ret = "Server Solutions Premium"}
            53 { $ret = "Server Solutions Premium (core installation)"}
            54 { $ret = "Server For SB Solutions EM"}
            55 { $ret = "Server For SB Solutions EM"}
            56 { $ret = "Windows MultiPoint Server"}
            59 { $ret = "Windows Essential Server Solution Management"}
            60 { $ret = "Windows Essential Server Solution Additional"}
            61 { $ret = "Windows Essential Server Solution Management SVC" }
            62 { $ret = "Windows Essential Server Solution Additional SVC" }
            63 { $ret = "Small Business Server Premium (core installation)" }
            64 { $ret = "Server Hyper Core V" }
            66 { $ret = "Starter E" }
            67 { $ret = "Home Basic E" }
            68 { $ret = "Home Premium E" }
            69 { $ret = "Professional E" }
            70 { $ret = "Enterprise E" }
            71 { $ret = "Ultimate E" }
            72 { $ret = "Server Enterprise (evaluation installation)" }
            76 { $ret = "Windows MultiPoint Server Standard (full installation)" }
            77 { $ret = "Windows MultiPoint Server Premium (full installation)" }
            79 { $ret = "Server Standard (evaluation installation)" }
            80 { $ret = "Server Datacenter (evaluation installation)" }
            84 { $ret = "Enterprise N (evaluation installation)" }
            95 { $ret = "Storage Server Workgroup (evaluation installation)" }
            96 { $ret = "Storage Server Standard (evaluation installation)" }
            98 { $ret = "Windows 8 N" }
            99 { $ret = "Windows 8 China" }
            100 { $ret = "Windows 8 Single Language" }
            101 { $ret = "Windows 8" }
            default { $ret = $intOperatingSystemSKU }
      }
      return $ret
}

function get_Win32_OperatingSystem_CountryCode($intCountryCode)
{
    switch ($intCountryCode )
    {
                  1 { $ret = "United States" }
                  7 { $ret = "Kazakhstan" }
                  7 { $ret = "Russia" }
                  20 { $ret = "Egypt" }
                  27 { $ret = "South Africa" }
                  30 { $ret = "Greece" }
                  31 { $ret = "Netherlands" }
                  32 { $ret = "Belgium" }
                  33 { $ret = "France" }
                  34 { $ret = "Spain" }
                  36 { $ret = "Hungary" }
                  39 { $ret = "Holy See (Vatican City)" }
                  39 { $ret = "Italy" }
                  40 { $ret = "Romania" }
                  41 { $ret = "Switzerland" }
                  43 { $ret = "Austria" }
                  44 { $ret = "United Kingdom" }
                  45 { $ret = "Denmark" }
                  46 { $ret = "Sweden" }
                  47 { $ret = "Norway" }
                  48 { $ret = "Poland" }
                  49 { $ret = "Germany" }
                  51 { $ret = "Peru" }
                  52 { $ret = "Mexico" }
                  53 { $ret = "Cuba" }
                  54 { $ret = "Argentina" }
                  55 { $ret = "Brazil" }
                  56 { $ret = "Chile" }
                  57 { $ret = "Colombia" }
                  58 { $ret = "Venezuela" }
                  60 { $ret = "Malaysia" }
                  61 { $ret = "Australia" }
                  61 { $ret = "Christmas Island" }
                  61 { $ret = "Cocos (Keeling) Islands" }
                  62 { $ret = "Indonesia" }
                  63 { $ret = "Philippines" }
                  64 { $ret = "New Zealand" }
                  65 { $ret = "Singapore" }
                  66 { $ret = "Thailand" }
                  81 { $ret = "Japan" }
                  82 { $ret = "South Korea" }
                  84 { $ret = "Vietnam" }
                  86 { $ret = "China" }
                  90 { $ret = "Turkey" }
                  91 { $ret = "India" }
                  92 { $ret = "Pakistan" }
                  93 { $ret = "Afghanistan" }
                  94 { $ret = "Sri Lanka" }
                  95 { $ret = "Burma (Myanmar)" }
                  98 { $ret = "Iran" }
                  212 { $ret = "Morocco" }
                  213 { $ret = "Algeria" }
                  216 { $ret = "Tunisia" }
                  218 { $ret = "Libya" }
                  220 { $ret = "Gambia" }
                  221 { $ret = "Senegal" }
                  222 { $ret = "Mauritania" }
                  223 { $ret = "Mali" }
                  224 { $ret = "Guinea" }
                  225 { $ret = "Ivory Coast" }
                  226 { $ret = "Burkina Faso" }
                  227 { $ret = "Niger" }
                  228 { $ret = "Togo" }
                  229 { $ret = "Benin" }
                  230 { $ret = "Mauritius" }
                  231 { $ret = "Liberia" }
                  232 { $ret = "Sierra Leone" }
                  233 { $ret = "Ghana" }
                  234 { $ret = "Nigeria" }
                  235 { $ret = "Chad" }
                  236 { $ret = "Central African Republic" }
                  237 { $ret = "Cameroon" }
                  238 { $ret = "Cape Verde" }
                  239 { $ret = "Sao Tome and Principe" }
                  240 { $ret = "Equatorial Guinea" }
                  241 { $ret = "Gabon" }
                  242 { $ret = "Republic of the Congo" }
                  243 { $ret = "Democratic Republic of the Congo" }
                  244 { $ret = "Angola" }
                  245 { $ret = "Guinea-Bissau" }
                  248 { $ret = "Seychelles" }
                  249 { $ret = "Sudan" }
                  250 { $ret = "Rwanda" }
                  251 { $ret = "Ethiopia" }
                  252 { $ret = "Somalia" }
                  253 { $ret = "Djibouti" }
                  254 { $ret = "Kenya" }
                  255 { $ret = "Tanzania" }
                  256 { $ret = "Uganda" }
                  257 { $ret = "Burundi" }
                  258 { $ret = "Mozambique" }
                  260 { $ret = "Zambia" }
                  261 { $ret = "Madagascar" }
                  262 { $ret = "Mayotte" }
                  263 { $ret = "Zimbabwe" }
                  264 { $ret = "Namibia" }
                  265 { $ret = "Malawi" }
                  266 { $ret = "Lesotho" }
                  267 { $ret = "Botswana" }
                  268 { $ret = "Swaziland" }
                  269 { $ret = "Comoros" }
                  290 { $ret = "Saint Helena" }
                  291 { $ret = "Eritrea" }
                  297 { $ret = "Aruba" }
                  298 { $ret = "Faroe Islands" }
                  299 { $ret = "Greenland" }
                  350 { $ret = "Gibraltar" }
                  351 { $ret = "Portugal" }
                  352 { $ret = "Luxembourg" }
                  353 { $ret = "Ireland" }
                  354 { $ret = "Iceland" }
                  355 { $ret = "Albania" }
                  356 { $ret = "Malta" }
                  357 { $ret = "Cyprus" }
                  358 { $ret = "Finland" }
                  359 { $ret = "Bulgaria" }
                  370 { $ret = "Lithuania" }
                  371 { $ret = "Latvia" }
                  372 { $ret = "Estonia" }
                  373 { $ret = "Moldova" }
                  374 { $ret = "Armenia" }
                  375 { $ret = "Belarus" }
                  376 { $ret = "Andorra" }
                  377 { $ret = "Monaco" }
                  378 { $ret = "San Marino" }
                  380 { $ret = "Ukraine" }
                  381 { $ret = "Kosovo" }
                  381 { $ret = "Serbia" }
                  382 { $ret = "Montenegro" }
                  385 { $ret = "Croatia" }
                  386 { $ret = "Slovenia" }
                  387 { $ret = "Bosnia and Herzegovina" }
                  389 { $ret = "Macedonia" }
                  420 { $ret = "Czech Republic" }
                  421 { $ret = "Slovakia" }
                  423 { $ret = "Liechtenstein" }
                  500 { $ret = "Falkland Islands" }
                  501 { $ret = "Belize" }
                  502 { $ret = "Guatemala" }
                  503 { $ret = "El Salvador" }
                  504 { $ret = "Honduras" }
                  505 { $ret = "Nicaragua" }
                  506 { $ret = "Costa Rica" }
                  507 { $ret = "Panama" }
                  508 { $ret = "Saint Pierre and Miquelon" }
                  509 { $ret = "Haiti" }
                  590 { $ret = "Saint Barthelemy" }
                  591 { $ret = "Bolivia" }
                  592 { $ret = "Guyana" }
                  593 { $ret = "Ecuador" }
                  595 { $ret = "Paraguay" }
                  597 { $ret = "Suriname" }
                  598 { $ret = "Uruguay" }
                  599 { $ret = "Netherlands Antilles" }
                  670 { $ret = "Timor-Leste" }
                  672 { $ret = "Antarctica" }
                  672 { $ret = "Norfolk Island" }
                  673 { $ret = "Brunei" }
                  674 { $ret = "Nauru" }
                  675 { $ret = "Papua New Guinea" }
                  676 { $ret = "Tonga" }
                  677 { $ret = "Solomon Islands" }
                  678 { $ret = "Vanuatu" }
                  679 { $ret = "Fiji" }
                  680 { $ret = "Palau" }
                  681 { $ret = "Wallis and Futuna" }
                  682 { $ret = "Cook Islands" }
                  683 { $ret = "Niue" }
                  685 { $ret = "Samoa" }
                  686 { $ret = "Kiribati" }
                  687 { $ret = "New Caledonia" }
                  688 { $ret = "Tuvalu" }
                  689 { $ret = "French Polynesia" }
                  690 { $ret = "Tokelau" }
                  691 { $ret = "Micronesia" }
                  692 { $ret = "Marshall Islands" }
                  850 { $ret = "North Korea" }
                  852 { $ret = "Hong Kong" }
                  853 { $ret = "Macau" }
                  855 { $ret = "Cambodia" }
                  856 { $ret = "Laos" }
                  870 { $ret = "Pitcairn Islands" }
                  880 { $ret = "Bangladesh" }
                  886 { $ret = "Taiwan" }
                  960 { $ret = "Maldives" }
                  961 { $ret = "Lebanon" }
                  962 { $ret = "Jordan" }
                  963 { $ret = "Syria" }
                  964 { $ret = "Iraq" }
                  965 { $ret = "Kuwait" }
                  966 { $ret = "Saudi Arabia" }
                  967 { $ret = "Yemen" }
                  968 { $ret = "Oman" }
                  970 { $ret = "Gaza Strip" }
                  970 { $ret = "West Bank" }
                  971 { $ret = "United Arab Emirates" }
                  972 { $ret = "Israel" }
                  973 { $ret = "Bahrain" }
                  974 { $ret = "Qatar" }
                  975 { $ret = "Bhutan" }
                  976 { $ret = "Mongolia" }
                  977 { $ret = "Nepal" }
                  992 { $ret = "Tajikistan" }
                  993 { $ret = "Turkmenistan" }
                  994 { $ret = "Azerbaijan" }
                  995 { $ret = "Georgia" }
                  996 { $ret = "Kyrgyzstan" }
                  998 { $ret = "Uzbekistan" }
                  1242 { $ret = "Bahamas" }
                  1246 { $ret = "Barbados" }
                  1264 { $ret = "Anguilla" }
                  1268 { $ret = "Antigua and Barbuda" }
                  1284 { $ret = "British Virgin Islands" }
                  1340 { $ret = "US Virgin Islands" }
                  1345 { $ret = "Cayman Islands" }
                  1441 { $ret = "Bermuda" }
                  1473 { $ret = "Grenada" }
                  1599 { $ret = "Saint Martin" }
                  1649 { $ret = "Turks and Caicos Islands" }
                  1664 { $ret = "Montserrat" }
                  1670 { $ret = "Northern Mariana Islands" }
                  1671 { $ret = "Guam" }
                  1684 { $ret = "American Samoa" }
                  1758 { $ret = "Saint Lucia" }
                  1767 { $ret = "Dominica" }
                  1784 { $ret = "Saint Vincent and the Grenadines" }
                  1809 { $ret = "Dominican Republic" }
                  1868 { $ret = "Trinidad and Tobago" }
                  1869 { $ret = "Saint Kitts and Nevis" }
                  1876 { $ret = "Jamaica" }
            default { $ret = $intCountryCode }
      }
      return $ret
}


function get_Win32_OperatingSystem_OSLanguage($intOSLanguage)
{
    switch ($intOSLanguage )
    {
            1  { $ret = "Arabic" }
            4  { $ret = "Chinese (Simplified)– China" }
            9  { $ret = "English" }
            1025  { $ret = "Arabic – Saudi Arabia" }
            1026  { $ret = "Bulgarian" }
            1027  { $ret = "Catalan" }
            1028  { $ret = "Chinese (Traditional) – Taiwan" }
            1029  { $ret = "Czech" }
            1030  { $ret = "Danish" }
            1031  { $ret = "German – Germany" }
            1032  { $ret = "Greek" }
            1033  { $ret = "English – United States" }
            1034  { $ret = "Spanish – Traditional Sort" }
            1035  { $ret = "Finnish" }
            1036  { $ret = "French – France" }
            1037  { $ret = "Hebrew" }
            1038  { $ret = "Hungarian" }
            1039  { $ret = "Icelandic" }
            1040  { $ret = "Italian – Italy" }
            1041  { $ret = "Japanese" }
            1042  { $ret = "Korean" }
            1043  { $ret = "Dutch – Netherlands" }
            1044  { $ret = "Norwegian – Bokmal" }
            1045  { $ret = "Polish" }
            1046  { $ret = "Portuguese – Brazil" }
            1047  { $ret = "Rhaeto-Romanic" }
            1048  { $ret = "Romanian" }
            1049  { $ret = "Russian" }
            1050  { $ret = "Croatian" }
            1051  { $ret = "Slovak" }
            1052  { $ret = "Albanian" }
            1053  { $ret = "Swedish" }
            1054  { $ret = "Thai" }
            1055  { $ret = "Turkish" }
            1056  { $ret = "Urdu" }
            1057  { $ret = "Indonesian" }
            1058  { $ret = "Ukrainian" }
            1059  { $ret = "Belarusian" }
            1060  { $ret = "Slovenian" }
            1061  { $ret = "Estonian" }
            1062  { $ret = "Latvian" }
            1063  { $ret = "Lithuanian" }
            1065  { $ret = "Persian" }
            1066  { $ret = "Vietnamese" }
            1069  { $ret = "Basque (Basque) – Basque" }
            1070  { $ret = "Serbian" }
            1071  { $ret = "Macedonian (FYROM)" }
            1072  { $ret = "Sutu" }
            1073  { $ret = "Tsonga" }
            1074  { $ret = "Tswana" }
            1076  { $ret = "Xhosa" }
            1077  { $ret = "Zulu" }
            1078  { $ret = "Afrikaans" }
            1080  { $ret = "Faeroese" }
            1081  { $ret = "Hindi" }
            1082  { $ret = "Maltese" }
            1084  { $ret = "Scottish Gaelic (United Kingdom)" }
            1085  { $ret = "Yiddish" }
            1086  { $ret = "Malay – Malaysia" }
            2049  { $ret = "Arabic – Iraq" }
            2052  { $ret = "Chinese (Simplified) – PRC" }
            2055  { $ret = "German – Switzerland" }
            2057  { $ret = "English – United Kingdom" }
            2058  { $ret = "Spanish – Mexico" }
            2060  { $ret = "French – Belgium" }
            2064  { $ret = "Italian – Switzerland" }
            2067  { $ret = "Dutch – Belgium" }
            2068  { $ret = "Norwegian – Nynorsk" }
            2070  { $ret = "Portuguese – Portugal" }
            2072  { $ret = "Romanian – Moldova" }
            2073  { $ret = "Russian – Moldova" }
            2074  { $ret = "Serbian – Latin" }
            2077  { $ret = "Swedish – Finland" }
            3073  { $ret = "Arabic – Egypt" }
            3076  { $ret = "Chinese (Traditional) – Hong Kong SAR" }
            3079  { $ret = "German – Austria" }
            3081  { $ret = "English – Australia" }
            3082  { $ret = "Spanish – International Sort" }
            3084  { $ret = "French – Canada" }
            3098  { $ret = "Serbian – Cyrillic" }
            4097  { $ret = "Arabic – Libya" }
            4100  { $ret = "Chinese (Simplified) – Singapore" }
            4103  { $ret = "German – Luxembourg" }
            4105  { $ret = "English – Canada" }
            4106  { $ret = "Spanish – Guatemala" }
            4108  { $ret = "French – Switzerland" }
            5121  { $ret = "Arabic – Algeria" }
            5127  { $ret = "German – Liechtenstein" }
            5129  { $ret = "English – New Zealand" }
            5130  { $ret = "Spanish – Costa Rica" }
            5132  { $ret = "French – Luxembourg" }
            6145  { $ret = "Arabic – Morocco" }
            6153  { $ret = "English – Ireland" }
            6154  { $ret = "Spanish – Panama" }
            7169  { $ret = "Arabic – Tunisia" }
            7177  { $ret = "English – South Africa" }
            7178  { $ret = "Spanish – Dominican Republic" }
            8193  { $ret = "Arabic – Oman" }
            8201  { $ret = "English – Jamaica" }
            8202  { $ret = "Spanish – Venezuela" }
            9217  { $ret = "Arabic – Yemen" }
            9226  { $ret = "Spanish – Colombia" }
            10241  { $ret = "Arabic – Syria" }
            10249  { $ret = "English – Belize" }
           10250  { $ret = "Spanish – Peru" }
            11265  { $ret = "Arabic – Jordan" }
            11273  { $ret = "English – Trinidad" }
            11274  { $ret = "Spanish – Argentina" }
            12289  { $ret = "Arabic – Lebanon" }
            12298  { $ret = "Spanish – Ecuador" }
            13313  { $ret = "Arabic – Kuwait" }
            13322  { $ret = "Spanish – Chile" }
            14337  { $ret = "Arabic – U.A.E." }
            14346  { $ret = "Spanish – Uruguay" }
            15361  { $ret = "Arabic – Bahrain" }
            15370  { $ret = "Spanish – Paraguay" }
            16385  { $ret = "Arabic – Qatar" }
            16394  { $ret = "Spanish – Bolivia" }
            17418  { $ret = "Spanish – El Salvador" }
            18442  { $ret = "Spanish – Honduras" }
            19466  { $ret = "Spanish – Nicaragua" }
            20490  { $ret = "Spanish – Puerto Rico" }
            default { $ret = $intOSLanguage }
      }
      return $ret
}
#endregion

#region Win32_PhysicalMemory Functions
function Get-Win32_PhysicalMemory-FormFactor($uint16_FormFactor)
{
      switch ($uint16_FormFactor)
      {
            0 { $ret = "Unknown" }
            1 { $ret = "Other" }
            2 { $ret = "SIP" }
            3 { $ret = "DIP" }
            4 { $ret = "ZIP" }
            5 { $ret = "SOJ" }
            6 { $ret = "Proprietary" }
            7 { $ret = "SIMM" }
            8 { $ret = "DIMM" }
            9 { $ret = "TSOP" }
            10 { $ret = "PGA" }
            11 { $ret = "RIMM" }
            12 { $ret = "SODIMM" }
            13 { $ret = "SRIMM" }
            14 { $ret = "SMD" }
            15 { $ret = "SSMP" }
            16 { $ret = "QFP" }
            17 { $ret = "TQFP" }
            18 { $ret = "SOIC" }
            19 { $ret = "LCC" }
            20 { $ret = "PLCC" }
            21 { $ret = "BGA" }
            22 { $ret = "FPBGA" }
            23 { $ret = "LGA" }
            default { $ret = $uint16_FormFactor }
      }
      return $ret
}
function Get-Win32_PhysicalMemory-InterleavePosition($uint32_InterleavePosition)
{
      switch ($uint32_InterleavePosition)
      {
            0 { $ret = "Noninterleaved" }
            1 { $ret = "First position" }
            2 { $ret = "Second position" }
            default { $ret = $uint32_InterleavePosition }
      }
      return $ret
}
function Get-Win32_PhysicalMemory-MemoryType($uint16_MemoryType)
{
      switch ($uint16_MemoryType)
      {
            0 { $ret = "Unknown" }
            1 { $ret = "Other" }
            2 { $ret = "DRAM" }
            3 { $ret = "Synchronous DRAM" }
            4 { $ret = "Cache DRAM" }
            5 { $ret = "EDO" }
            6 { $ret = "EDRAM" }
            7 { $ret = "VRAM" }
            8 { $ret = "SRAM" }
            9 { $ret = "RAM" }
            10 { $ret = "ROM" }
            11 { $ret = "Flash" }
            12 { $ret = "EEPROM" }
            13 { $ret = "FEPROM" }
            14 { $ret = "EPROM" }
            15 { $ret = "CDRAM" }
            16 { $ret = "3DRAM" }
            17 { $ret = "SDRAM" }
            18 { $ret = "SGRAM" }
            19 { $ret = "RDRAM" }
            20 { $ret = "DDR" }
            21 { $ret = "DDR-2" }
            default { $ret = $uint16_MemoryType }
      }
      return $ret
}
#endregion Win32_PhysicalMemory Functions

#region Win32_Processor Functions
Function Get-Win32_Processor-Architecture($uint16_Architecture)
{
      switch ($uint16_Architecture)
      {
            0 { $ret = "x86" }
            1 { $ret = "MIPS" }
            2 { $ret = "Alpha" }
            3 { $ret = "PowerPC" }
            5 { $ret = "ARM" }
            6 { $ret = "Itanium-based systems" }
            9 { $ret = "x64" }
            default { $ret = $uint16_Architecture }
      }
      return $ret
}
Function Get-Win32_Processor-Availability($uint16_Availability)
{
      switch ($uint16_Availability)
      {
            1 { $ret = "Other" }
            2 { $ret = "Unknown" }
            3 { $ret = "Running or Full Power" }
            4 { $ret = "Warning" }
            5 { $ret = "In Test" }
            6 { $ret = "Not Applicable" }
            7 { $ret = "Power Off" }
            8 { $ret = "Off Line" }
            9 { $ret = "Off Duty" }
            10 { $ret = "Degraded" }
            11 { $ret = "Not Installed" }
            12 { $ret = "Install Error" }
            13 { $ret = "Power Save - Unknown" }
            14 { $ret = "Power Save - Low Power Mode" }
            15 { $ret = "Power Save - Standby" }
            16 { $ret = "Power Cycle" }
            17 { $ret = "Power Save - Warning" }
            default { $ret = $uint16_Availability }
      }
      return $ret
}
Function Get-Win32_Processor-ConfigManagerErrorCode($uint32_ConfigManagerErrorCode)
{
      switch ($uint32_ConfigManagerErrorCode)
      {
            0 { $ret = "Device is working properly." }
            1 { $ret = "Device is not configured correctly." }
            2 { $ret = "Windows cannot load the driver for this device." }
            3 { $ret = "Driver for this device might be corrupted or the  system may be  low on memory or other resources." }
            4 { $ret = "Device is not working properly. One of its drivers or the  registry might be corrupted." }
            5 { $ret = "Driver for the device requires a resource that Windows cannot manage." }
            6 { $ret = "Boot configuration for the device conflicts with other devices." }
            7 { $ret = "Cannot filter." }
            8 { $ret = "Driver loader for the device is missing." }
            9 { $ret = "Device is not working properly.  The controlling firmware is incorrectly reporting the resources for the device." }
            10 { $ret = "Device cannot start." }
            11 { $ret = "Device failed." }
            12 { $ret = "Device cannot find enough free resources to use." }
            13 { $ret = "Windows cannot verify the device's resources." }
            14 { $ret = "Device cannot work properly until the computer is restarted." }
            15 { $ret = "Device is not working properly due to a possible re-enumeration problem." }
            16 { $ret = "Windows cannot identify all of the resources that the device uses." }
            17 { $ret = "Device is requesting  an unknown resource type." }
            18 { $ret = "Device drivers must  be reinstalled." }
            19 { $ret = "Failure using the VxD loader." }
            20 { $ret = "Registry might be corrupted." }
            21 { $ret = "System failure. If changing the device driver is ineffective, see the hardware documentation. Windows is removing the device." }
            22 { $ret = "Device is disabled." }
            23 { $ret = "System failure. If changing the device driver is ineffective, see the hardware documentation." }
            24 { $ret = "Device is not present,  not working properly, or does not have all of its drivers installed." }
            25 { $ret = "Windows is still setting up the device." }
            26 { $ret = "Windows is still setting up the device." }
            27 { $ret = "Device does not have valid log configuration." }
            28 { $ret = "Device drivers   are not installed." }
            29 { $ret = "Device is disabled.  The device firmware   did not provide  the required resources." }
            30 { $ret = "Device is using an IRQ resource that another device is using." }
            31 { $ret = "Device is not working properly.  Windows cannot load the  required device drivers." }
            default { $ret = $uint32_ConfigManagerErrorCode }
      }
      return $ret
}
Function Get-Win32_Processor-CpuStatus($uint16_CpuStatus)
{
      switch ($uint16_CpuStatus)
      {
            0 { $ret = "Unknown" }
            1 { $ret = "CPU Enabled" }
            2 { $ret = "CPU Disabled by User via BIOS Setup" }
            3 { $ret = "CPU Disabled by BIOS (POST Error)" }
            4 { $ret = "CPU Is Idle" }
            5 { $ret = "Reserved" }
            6 { $ret = "Reserved" }
            7 { $ret = "Other" }
            default { $ret = $uint16_CpuStatus }
      }
      return $ret
}
Function Get-Win32_Processor-Family($uint16_Family)
{
      switch ($uint16_Family)
      {
            1 { $ret = "Other" }
            2 { $ret = "Unknown" }
            3 { $ret = "8086" }
            4 { $ret = "80286" }
            5 { $ret = "Intel386™ Processor" }
            6 { $ret = "Intel486™ Processor" }
            7 { $ret = "8087" }
            8 { $ret = "80287" }
            9 { $ret = "80387" }
            10 { $ret = "80487" }
            11 { $ret = "Pentium Brand" }
            12 { $ret = "Pentium Pro" }
            13 { $ret = "Pentium II" }
            14 { $ret = "Pentium Processor with MMX™ Technology" }
            15 { $ret = "Celeron™" }
            16 { $ret = "Pentium II Xeon™" }
            17 { $ret = "Pentium III" }
            18 { $ret = "M1 Family" }
            19 { $ret = "M2 Family" }
            24 { $ret = "AMD Duron™ Processor Family" }
            25 { $ret = "K5 Family" }
            26 { $ret = "K6 Family" }
            27 { $ret = "K6-2" }
            28 { $ret = "K6-3" }
            29 { $ret = "AMD Athlon™ Processor Family" }
            30 { $ret = "AMD2900 Family" }
            31 { $ret = "K6-2+" }
            32 { $ret = "Power PC Family" }
            33 { $ret = "Power PC 601" }
            34 { $ret = "Power PC 603" }
            35 { $ret = "Power PC 603+" }
            36 { $ret = "Power PC 604" }
            37 { $ret = "Power PC 620" }
            38 { $ret = "Power PC X704" }
            39 { $ret = "Power PC 750" }
            48 { $ret = "Alpha Family" }
            49 { $ret = "Alpha 21064" }
            50 { $ret = "Alpha 21066" }
            51 { $ret = "Alpha 21164" }
            52 { $ret = "Alpha 21164PC" }
            53 { $ret = "Alpha 21164a" }
            54 { $ret = "Alpha 21264" }
            55 { $ret = "Alpha 21364" }
            64 { $ret = "MIPS Family" }
            65 { $ret = "MIPS R4000" }
            66 { $ret = "MIPS R4200" }
            67 { $ret = "MIPS R4400" }
            68 { $ret = "MIPS R4600" }
            69 { $ret = "MIPS R10000" }
            80 { $ret = "SPARC Family" }
            81 { $ret = "SuperSPARC" }
            82 { $ret = "microSPARC II" }
            83 { $ret = "microSPARC IIep" }
            84 { $ret = "UltraSPARC" }
            85 { $ret = "UltraSPARC II" }
            86 { $ret = "UltraSPARC IIi" }
            87 { $ret = "UltraSPARC III" }
            88 { $ret = "UltraSPARC IIIi" }
            96 { $ret = "68040" }
            97 { $ret = "68xxx Family" }
            98 { $ret = "68000" }
            99 { $ret = "68010" }
            100 { $ret = "68020" }
            101 { $ret = "68030" }
            112 { $ret = "Hobbit Family" }
            120 { $ret = "Crusoe™ TM5000 Family" }
            121 { $ret = "Crusoe™ TM3000 Family" }
            122 { $ret = "Efficeon™ TM8000 Family" }
            128 { $ret = "Weitek" }
            130 { $ret = "Itanium™ Processor" }
            131 { $ret = "AMD Athlon™ 64 Processor Family" }
            132 { $ret = "AMD Opteron™ Processor Family" }
            144 { $ret = " PA-RISC Family" }
            145 { $ret = "PA-RISC 8500" }
            146 { $ret = "PA-RISC 8000" }
            147 { $ret = "PA-RISC 7300LC" }
            148 { $ret = "PA-RISC 7200" }
            149 { $ret = "PA-RISC 7100LC" }
            150 { $ret = "PA-RISC 7100" }
            160 { $ret = "V30 Family" }
            176 { $ret = "Pentium III Xeon™ Processor" }
            177 { $ret = "Pentium III Processor with Intel SpeedStep™ Technology" }
            178 { $ret = "Pentium 4" }
            179 { $ret = "Intel Xeon™" }
            180 { $ret = "AS400 Family" }
            181 { $ret = "Intel Xeon™ Processor MP" }
            182 { $ret = "AMD Athlon™ XP Family" }
            183 { $ret = "AMD Athlon™ MP Family" }
            184 { $ret = "Intel Itanium 2" }
            185 { $ret = "Intel Pentium M Processor" }
            190 { $ret = "K7" }
            200 { $ret = "IBM390 Family" }
            201 { $ret = "G4" }
            202 { $ret = "G5" }
            203 { $ret = "G6" }
            204 { $ret = "z/Architecture Base" }
            250 { $ret = "i860" }
            251 { $ret = "i960" }
            260 { $ret = "SH-3" }
            261 { $ret = "SH-4" }
            280 { $ret = "ARM" }
            281 { $ret = "StrongARM" }
            300 { $ret = "6x86" }
            301 { $ret = "MediaGX" }
            302 { $ret = "MII" }
            320 { $ret = "WinChip" }
            350 { $ret = "DSP" }
            500 { $ret = "Video Processor" }
            default { $ret = $uint16_Family }
      }
      return $ret
}
Function Get-Win32_Processor-PowerManagementCapabilities($uint16_array_PowerManagementCapabilities)
{
      switch ($uint16_array_PowerManagementCapabilities)
      {
            0 { $ret = "Unknown" }
            1 { $ret = "Not Supported" }
            2 { $ret = "Disabled" }
            3 { $ret = "Enabled" }
            4 { $ret = "Power Saving Modes Entered Automatically" }
            5 { $ret = "Power State Settable" }
            6 { $ret = "Power Cycling Supported" }
            7 { $ret = "Timed Power-On Supported" }
            default { $ret = $uint16_array_PowerManagementCapabilities }
      }
      return $ret
}
Function Get-Win32_Processor-ProcessorType($uint16_ProcessorType)
{
      switch ($uint16_ProcessorType)
      {
            1 { $ret = "Other" }
            2 { $ret = "Unknown" }
            3 { $ret = "Central Processor" }
            4 { $ret = "Math Processor" }
            5 { $ret = "DSP Processor" }
            6 { $ret = "Video Processor" }
            default { $ret = $uint16_ProcessorType }
      }
      return $ret
}
Function Get-Win32_Processor-StatusInfo($uint16_StatusInfo)
{
      switch ($uint16_StatusInfo)
      {
            1 { $ret = "Other" }
            2 { $ret = "Unknown" }
            3 { $ret = "Enabled" }
            4 { $ret = "Disabled" }
            5 { $ret = "Not Applicable" }
            default { $ret = $uint16_StatusInfo }
      }
      return $ret
}
Function Get-Win32_Processor-UpgradeMethod($uint16_UpgradeMethod)
{
      switch ($uint16_UpgradeMethod)
      {
            1 { $ret = "Other" }
            2 { $ret = "Unknown" }
            3 { $ret = "Daughter Board" }
            4 { $ret = "ZIF Socket" }
            5 { $ret = "Replacement or Piggy Back" }
            6 { $ret = "None" }
            7 { $ret = "LIF Socket" }
            8 { $ret = "Slot 1" }
            9 { $ret = "Slot 2" }
            10 { $ret = "370 Pin Socket" }
            11 { $ret = "Slot A" }
            12 { $ret = "Slot M" }
            13 { $ret = "Socket 423" }
            14 { $ret = "Socket A (Socket 462)" }
            15 { $ret = "Socket 478" }
            16 { $ret = "Socket 754" }
            17 { $ret = "Socket 940" }
            18 { $ret = "Socket 939" }
            default { $ret = $uint16_UpgradeMethod }
      }
      return $ret
}
#endregion Win32_Processor Functions

function get_CountryCode($strCountryCode)
{
    switch ($strCountryCode)
    {
        "AF" { $ret = "Afghanistan" }
        "AL" { $ret = "Albania" }
        "DZ" { $ret = "Algeria" }
        "AS" { $ret = "American Samoa" } 
        "AD" { $ret = "Andorra" }
        "AO" { $ret = "Angola" } 
        "AI" { $ret = "Anguilla" } 
        "AQ" { $ret = "Antarctica" } 
        "AG" { $ret = "Antigua And Barbuda" }
        "AR" { $ret = "Argentina" }
        "AM" { $ret = "Armenia" }
        "AW" { $ret = "Aruba" }
        "AU" { $ret = "Australia" }
        "AT" { $ret = "Austria" }
        "AZ" { $ret = "Azerbaijan" } 
        "BS" { $ret = "Bahamas" }
        "BH" { $ret = "Bahrain" }
        "BD" { $ret = "Bangladesh" } 
        "BB" { $ret = "Barbados" } 
        "BY" { $ret = "Belarus" }
        "BE" { $ret = "Belgium" }
        "BZ" { $ret = "Belize" } 
        "BJ" { $ret = "Benin" }
        "BM" { $ret = "Bermuda" }
        "BT" { $ret = "Bhutan" } 
        "BO" { $ret = "Bolivia" }
        "BA" { $ret = "Bosnia And Herzegovina" } 
        "BW" { $ret = "Botswana" } 
        "BV" { $ret = "Bouvet Island" }
        "BR" { $ret = "Brazil" } 
        "IO" { $ret = "British Indian Ocean Territory" } 
        "BN" { $ret = "Brunei Darussalam" }
        "BG" { $ret = "Bulgaria" } 
        "BF" { $ret = "Burkina Faso" } 
        "BI" { $ret = "Burundi" }
        "KH" { $ret = "Cambodia" } 
        "CM" { $ret = "Cameroon" } 
        "CA" { $ret = "Canada" } 
        "CV" { $ret = "Cape Verde" } 
        "KY" { $ret = "Cayman Islands" } 
        "CF" { $ret = "Central African Republic" } 
        "TD" { $ret = "Chad" } 
        "CL" { $ret = "Chile" }
        "CN" { $ret = "China" }
        "CX" { $ret = "Christmas Island" } 
        "CC" { $ret = "Cocos (keeling) Islands" }
       "CO" { $ret = "Colombia" } 
        "KM" { $ret = "Comoros" }
        "CG" { $ret = "Congo" }
        "CD" { $ret = "Congo, The Democratic Republic Of The" }
        "CK" { $ret = "Cook Islands" } 
        "CR" { $ret = "Costa Rica" } 
        "CI" { $ret = "Cote D'ivoire" }
        "HR" { $ret = "Croatia" }
        "CU" { $ret = "Cuba" } 
        "CY" { $ret = "Cyprus" } 
        "CZ" { $ret = "Czech Republic" } 
        "DK" { $ret = "Denmark" }
        "DJ" { $ret = "Djibouti" } 
        "DM" { $ret = "Dominica" } 
        "DO" { $ret = "Dominican Republic" } 
        "TP" { $ret = "East Timor" } 
        "EC" { $ret = "Ecuador" }
        "EG" { $ret = "Egypt" }
        "SV" { $ret = "El Salvador" }
        "GQ" { $ret = "Equatorial Guinea" }
        "ER" { $ret = "Eritrea" }
        "EE" { $ret = "Estonia" }
        "ET" { $ret = "Ethiopia" } 
        "FK" { $ret = "Falkland Islands (malvinas)" }
        "FO" { $ret = "Faroe Islands" }
        "FJ" { $ret = "Fiji" } 
        "FI" { $ret = "Finland" }
        "FR" { $ret = "France" } 
        "GF" { $ret = "French Guiana" }
        "PF" { $ret = "French Polynesia" } 
        "TF" { $ret = "French Southern Territories" }
        "GA" { $ret = "Gabon" }
        "GM" { $ret = "Gambia" } 
        "GE" { $ret = "Georgia" }
        "DE" { $ret = "Germany" }
        "GH" { $ret = "Ghana" }
        "GI" { $ret = "Gibraltar" }
        "GR" { $ret = "Greece" } 
        "GL" { $ret = "Greenland" }
        "GD" { $ret = "Grenada" }
        "GP" { $ret = "Guadeloupe" } 
        "GU" { $ret = "Guam" } 
        "GT" { $ret = "Guatemala" }
        "GN" { $ret = "Guinea" } 
        "GW" { $ret = "Guinea-bissau" }
        "GY" { $ret = "Guyana" } 
        "HT" { $ret = "Haiti" }
        "HM" { $ret = "Heard Island And Mcdonald Islands" }
        "VA" { $ret = "Holy See (vatican City State)" }
        "HN" { $ret = "Honduras" } 
        "HK" { $ret = "Hong Kong" }
        "HU" { $ret = "Hungary" }
        "IS" { $ret = "Iceland" }
        "IN" { $ret = "India" }
        "ID" { $ret = "Indonesia" }
        "IR" { $ret = "Iran, Islamic Republic Of" }
        "IQ" { $ret = "Iraq" } 
        "IE" { $ret = "Ireland" }
        "IL" { $ret = "Israel" } 
        "IT" { $ret = "Italy" }
        "JM" { $ret = "Jamaica" }
        "JP" { $ret = "Japan" }
        "JO" { $ret = "Jordan" } 
        "KZ" { $ret = "Kazakstan" }
        "KE" { $ret = "Kenya" }
        "KI" { $ret = "Kiribati" } 
        "KP" { $ret = "Korea, Democratic People's Republic Of" } 
        "KR" { $ret = "Korea, Republic Of" } 
        "KV" { $ret = "Kosovo" } 
        "KW" { $ret = "Kuwait" } 
        "KG" { $ret = "Kyrgyzstan" } 
        "LA" { $ret = "Lao People's Democratic Republic" } 
        "LV" { $ret = "Latvia" } 
        "LB" { $ret = "Lebanon" }
        "LS" { $ret = "Lesotho" }
        "LR" { $ret = "Liberia" }
        "LY" { $ret = "Libyan Arab Jamahiriya" } 
        "LI" { $ret = "Liechtenstein" }
        "LT" { $ret = "Lithuania" }
        "LU" { $ret = "Luxembourg" } 
        "MO" { $ret = "Macau" }
        "MK" { $ret = "Macedonia, The Former Yugoslav Republic Of" } 
        "MG" { $ret = "Madagascar" } 
        "MW" { $ret = "Malawi" } 
        "MY" { $ret = "Malaysia" } 
        "MV" { $ret = "Maldives" } 
        "ML" { $ret = "Mali" } 
        "MT" { $ret = "Malta" }
        "MH" { $ret = "Marshall Islands" } 
        "MQ" { $ret = "Martinique" } 
        "MR" { $ret = "Mauritania" } 
        "MU" { $ret = "Mauritius" }
        "YT" { $ret = "Mayotte" }
        "MX" { $ret = "Mexico" } 
        "FM" { $ret = "Micronesia, Federated States Of" }
        "MD" { $ret = "Moldova, Republic Of" } 
        "MC" { $ret = "Monaco" } 
        "MN" { $ret = "Mongolia" } 
        "MS" { $ret = "Montserrat" } 
        "ME" { $ret = "Montenegro" } 
        "MA" { $ret = "Morocco" }
        "MZ" { $ret = "Mozambique" } 
        "MM" { $ret = "Myanmar" }
        "NA" { $ret = "Namibia" }
        "NR" { $ret = "Nauru" }
        "NP" { $ret = "Nepal" }
        "NL" { $ret = "Netherlands" }
        "AN" { $ret = "Netherlands Antilles" } 
        "NC" { $ret = "New Caledonia" }
        "NZ" { $ret = "New Zealand" }
        "NI" { $ret = "Nicaragua" }
        "NE" { $ret = "Niger" }
        "NG" { $ret = "Nigeria" }
        "NU" { $ret = "Niue" } 
        "NF" { $ret = "Norfolk Island" } 
        "MP" { $ret = "Northern Mariana Islands" } 
        "NO" { $ret = "Norway" } 
        "OM" { $ret = "Oman" } 
        "PK" { $ret = "Pakistan" } 
        "PW" { $ret = "Palau" }
        "PS" { $ret = "Palestinian Territory, Occupied" }
        "PA" { $ret = "Panama" } 
        "PG" { $ret = "Papua New Guinea" } 
        "PY" { $ret = "Paraguay" } 
        "PE" { $ret = "Peru" } 
        "PH" { $ret = "Philippines" }
        "PN" { $ret = "Pitcairn" } 
        "PL" { $ret = "Poland" } 
        "PT" { $ret = "Portugal" } 
        "PR" { $ret = "Puerto Rico" }
        "QA" { $ret = "Qatar" }
        "RE" { $ret = "Reunion" }
        "RO" { $ret = "Romania" }
        "RU" { $ret = "Russian Federation" } 
        "RW" { $ret = "Rwanda" } 
        "SH" { $ret = "Saint Helena" } 
        "KN" { $ret = "Saint Kitts And Nevis" }
        "LC" { $ret = "Saint Lucia" }
        "PM" { $ret = "Saint Pierre And Miquelon" }
        "VC" { $ret = "Saint Vincent And The Grenadines" } 
        "WS" { $ret = "Samoa" }
        "SM" { $ret = "San Marino" } 
        "ST" { $ret = "Sao Tome And Principe" }
        "SA" { $ret = "Saudi Arabia" } 
        "SN" { $ret = "Senegal" }
        "RS" { $ret = "Serbia" } 
        "SC" { $ret = "Seychelles" } 
        "SL" { $ret = "Sierra Leone" } 
        "SG" { $ret = "Singapore" }
        "SK" { $ret = "Slovakia" } 
        "SI" { $ret = "Slovenia" } 
        "SB" { $ret = "Solomon Islands" }
        "SO" { $ret = "Somalia" }
        "ZA" { $ret = "South Africa" } 
        "GS" { $ret = "South Georgia And The South Sandwich Islands" } 
        "ES" { $ret = "Spain" }
        "LK" { $ret = "Sri Lanka" }
        "SD" { $ret = "Sudan" }
        "SR" { $ret = "Suriname" } 
        "SJ" { $ret = "Svalbard And Jan Mayen" } 
        "SZ" { $ret = "Swaziland" }
        "SE" { $ret = "Sweden" } 
        "CH" { $ret = "Switzerland" }
        "SY" { $ret = "Syrian Arab Republic" } 
        "TW" { $ret = "Taiwan, Province Of China" }
        "TJ" { $ret = "Tajikistan" } 
        "TZ" { $ret = "Tanzania, United Republic Of" } 
        "TH" { $ret = "Thailand" } 
        "TG" { $ret = "Togo" } 
        "TK" { $ret = "Tokelau" }
        "TO" { $ret = "Tonga" }
        "TT" { $ret = "Trinidad And Tobago" }
        "TN" { $ret = "Tunisia" }
        "TR" { $ret = "Turkey" } 
        "TM" { $ret = "Turkmenistan" } 
        "TC" { $ret = "Turks And Caicos Islands" } 
        "TV" { $ret = "Tuvalu" } 
        "UG" { $ret = "Uganda" } 
        "UA" { $ret = "Ukraine" }
        "AE" { $ret = "United Arab Emirates" } 
        "GB" { $ret = "United Kingdom" } 
        "US" { $ret = "United States" }
        "UM" { $ret = "United States Minor Outlying Islands" } 
        "UY" { $ret = "Uruguay" }
        "UZ" { $ret = "Uzbekistan" } 
        "VU" { $ret = "Vanuatu" }
        "VE" { $ret = "Venezuela" }
        "VN" { $ret = "Viet Nam" } 
        "VG" { $ret = "Virgin Islands, British" }
        "VI" { $ret = "Virgin Islands, U.s." } 
        "WF" { $ret = "Wallis And Futuna" }
        "EH" { $ret = "Western Sahara" } 
        "YE" { $ret = "Yemen" }
        "ZM" { $ret = "Zambia" } 
        "ZW" { $ret = "Zimbabwe" } 
    }
    return $ret
}
function get_CountryCodebyNumber($intCountryCode)
{
    switch ($intCountryCode)
    {
        "1" { $ret = "United States" }
        "7" { $ret = "Russia" }
        "20" { $ret = "Egypt" }
        "27" { $ret = "South Africa" }
        "30" { $ret = "Greece" }
        "31" { $ret = "Netherlands, The" }
        "32" { $ret = "Belgium" }
        "33" { $ret = "France" }
        "34" { $ret = "Spain" }
        "36" { $ret = "Hungary" }
        "39" { $ret = "Italy" }
        "40" { $ret = "Romania" }
        "41" { $ret = "Switzerland" }
        "43" { $ret = "Austria" }
        "44" { $ret = "United Kingdom" }
        "45" { $ret = "Denmark" }
        "46" { $ret = "Sweden" }
        "47" { $ret = "Norway" }
        "48" { $ret = "Poland" }
        "49" { $ret = "Germany" }
        "51" { $ret = "Peru" }
        "52" { $ret = "Mexico" }
        "53" { $ret = "Cuba" }
        "54" { $ret = "Argentina" }
        "55" { $ret = "Brazil" }
        "56" { $ret = "Chile" }
        "57" { $ret = "Colombia" }
        "58" { $ret = "Venezuela" }
        "60" { $ret = "Malaysia" }
        "61" { $ret = "Australia" }
        "62" { $ret = "Indonesia" }
        "63" { $ret = "Philippines" }
        "64" { $ret = "New Zealand" }
        "65" { $ret = "Singapore" }
        "66" { $ret = "Thailand" }
        "81" { $ret = "Japan" }
        "82" { $ret = "Korea" }
        "84" { $ret = "Viet Nam" }
        "86" { $ret = "China" }
        "90" { $ret = "Turkey" }
        "91" { $ret = "India" }
        "92" { $ret = "Pakistan" }
        "93" { $ret = "Afghanistan" }
        "94" { $ret = "Sri Lanka" }
        "95" { $ret = "Myanmar" }
        "98" { $ret = "Iran" }
        "101" { $ret = "Anguilla" }
        "103" { $ret = "Bahamas, The" }
        "104" { $ret = "Barbados" }
        "105" { $ret = "Bermuda" }
        "106" { $ret = "Virgin Islands, British" }
        "107" { $ret = "Canada" }
        "108" { $ret = "Cayman Islands" }
        "109" { $ret = "Dominica" }
        "110" { $ret = "Dominican Republic" }
        "111" { $ret = "Grenada" }
        "112" { $ret = "Jamaica" }
        "113" { $ret = "Montserrat" }
        "115" { $ret = "St. Kitts and Nevis" }
        "116" { $ret = "St. Vincent and the Grenadines" }
        "117" { $ret = "Trinidad and Tobago" }
        "118" { $ret = "Turks and Caicos Islands" }
        "121" { $ret = "Puerto Rico" }
        "122" { $ret = "St. Lucia" }
        "123" { $ret = "Virgin Islands" }
        "124" { $ret = "Guam" }
        "212" { $ret = "Morocco" }
        "213" { $ret = "Algeria" }
        "216" { $ret = "Tunisia" }
        "218" { $ret = "Libya" }
        "220" { $ret = "Gambia, The" }
        "221" { $ret = "Senegal" }
        "222" { $ret = "Mauritania" }
        "223" { $ret = "Mali" }
        "224" { $ret = "Guinea" }
        "225" { $ret = "Côte d'Ivoire" }
        "226" { $ret = "Burkina Faso" }
        "227" { $ret = "Niger" }
        "228" { $ret = "Togo" }
        "229" { $ret = "Benin" }
        "230" { $ret = "Mauritius" }
        "231" { $ret = "Liberia" }
        "232" { $ret = "Sierra Leone" }
        "233" { $ret = "Ghana" }
        "234" { $ret = "Nigeria" }
        "235" { $ret = "Chad" }
        "236" { $ret = "Central African Republic" }
        "237" { $ret = "Cameroon" }
        "238" { $ret = "Cape Verde" }
        "240" { $ret = "Equatorial Guinea" }
        "241" { $ret = "Gabon" }
        "242" { $ret = "Congo" }
        "243" { $ret = "Congo (DRC)" }
        "244" { $ret = "Angola" }
        "245" { $ret = "Guinea-Bissau" }
        "246" { $ret = "Diego Garcia" }
        "247" { $ret = "Ascension Island" }
        "248" { $ret = "Seychelles" }
        "249" { $ret = "Sudan" }
        "250" { $ret = "Rwanda" }
        "251" { $ret = "Ethiopia" }
        "252" { $ret = "Somalia" }
        "253" { $ret = "Djibouti" }
        "254" { $ret = "Kenya" }
        "255" { $ret = "Tanzania" }
        "256" { $ret = "Uganda" }
        "257" { $ret = "Burundi" }
        "258" { $ret = "Mozambique" }
        "260" { $ret = "Zambia" }
        "261" { $ret = "Madagascar" }
        "262" { $ret = "Reunion" }
        "263" { $ret = "Zimbabwe" }
        "264" { $ret = "Namibia" }
        "265" { $ret = "Malawi" }
        "266" { $ret = "Lesotho" }
        "267" { $ret = "Botswana" }
        "268" { $ret = "Swaziland" }
        "269" { $ret = "Mayotte" }
        "290" { $ret = "St. Helena" }
        "291" { $ret = "Eritrea" }
        "297" { $ret = "Aruba" }
        "298" { $ret = "Faroe Islands" }
        "299" { $ret = "Greenland" }
        "350" { $ret = "Gibraltar" }
        "351" { $ret = "Portugal" }
        "352" { $ret = "Luxembourg" }
        "353" { $ret = "Ireland" }
        "354" { $ret = "Iceland" }
        "355" { $ret = "Albania" }
        "356" { $ret = "Malta" }
        "357" { $ret = "Cyprus" }
        "358" { $ret = "Finland" }
        "359" { $ret = "Bulgaria" }
        "370" { $ret = "Lithuania" }
        "371" { $ret = "Latvia" }
        "372" { $ret = "Estonia" }
        "373" { $ret = "Moldova" }
        "374" { $ret = "Armenia" }
        "375" { $ret = "Belarus" }
        "376" { $ret = "Andorra" }
        "377" { $ret = "Monaco" }
        "378" { $ret = "San Marino" }
        "379" { $ret = "Vatican City" }
        "380" { $ret = "Ukraine" }
        "381" { $ret = "Yugoslavia" }
        "385" { $ret = "Croatia" }
        "386" { $ret = "Slovenia" }
        "387" { $ret = "Bosnia and Herzegovina" }
        "389" { $ret = "Macedonia, Former Yugoslav Republic of" }
        "420" { $ret = "Czech Republic" }
        "421" { $ret = "Slovakia" }
        "423" { $ret = "Liechtenstein" }
        "500" { $ret = "Falkland Islands (Islas Malvinas)" }
        "501" { $ret = "Belize" }
        "502" { $ret = "Guatemala" }
        "503" { $ret = "El Salvador" }
        "504" { $ret = "Honduras" }
        "505" { $ret = "Nicaragua" }
        "506" { $ret = "Costa Rica" }
        "507" { $ret = "Panama" }
        "508" { $ret = "St. Pierre and Miquelon" }
        "509" { $ret = "Haiti" }
        "590" { $ret = "Guadeloupe" }
        "591" { $ret = "Bolivia" }
        "592" { $ret = "Guyana" }
        "593" { $ret = "Ecuador" }
        "594" { $ret = "French Guiana" }
        "595" { $ret = "Paraguay" }
        "596" { $ret = "Martinique" }
        "597" { $ret = "Suriname" }
        "598" { $ret = "Uruguay" }
        "599" { $ret = "Netherlands Antilles" }
        "670" { $ret = "East Timor" }
        "672" { $ret = "Norfolk Island" }
        "673" { $ret = "Brunei" }
        "674" { $ret = "Nauru" }
        "675" { $ret = "Papua New Guinea" }
        "676" { $ret = "Tonga" }
        "677" { $ret = "Solomon Islands" }
        "678" { $ret = "Vanuatu" }
        "679" { $ret = "Fiji Islands" }
        "680" { $ret = "Palau" }
        "681" { $ret = "Wallis and Futuna" }
        "682" { $ret = "Cook Islands" }
        "683" { $ret = "Niue" }
        "684" { $ret = "American Samoa" }
        "685" { $ret = "Samoa" }
        "686" { $ret = "Kiribati" }
        "687" { $ret = "New Caledonia" }
        "688" { $ret = "Tuvalu" }
        "689" { $ret = "French Polynesia" }
        "690" { $ret = "Tokelau" }
        "691" { $ret = "Micronesia" }
        "692" { $ret = "Marshall Islands" }
        "705" { $ret = "Kazakhstan" }
        "850" { $ret = "North Korea" }
        "852" { $ret = "Hong Kong SAR" }
        "853" { $ret = "Macau SAR" }
        "855" { $ret = "Cambodia" }
        "856" { $ret = "Laos" }
        "880" { $ret = "Bangladesh" }
        "886" { $ret = "Taiwan" }
        "960" { $ret = "Maldives" }
        "961" { $ret = "Lebanon" }
        "962" { $ret = "Jordan" }
        "963" { $ret = "Syria" }
        "964" { $ret = "Iraq" }
        "965" { $ret = "Kuwait" }
        "966" { $ret = "Saudi Arabia" }
        "967" { $ret = "Yemen" }
        "968" { $ret = "Oman" }
        "971" { $ret = "United Arab Emirates" }
        "972" { $ret = "Israel" }
        "973" { $ret = "Bahrain" }
        "974" { $ret = "Qatar" }
        "975" { $ret = "Bhutan" }
        "976" { $ret = "Mongolia" }
        "977" { $ret = "Nepal" }
        "992" { $ret = "Tajikistan" }
        "993" { $ret = "Turkmenistan" }
        "994" { $ret = "Azerbaijan" }
        "995" { $ret = "Georgia" }
        "996" { $ret = "Kyrgyzstan" }
        "998" { $ret = "Uzbekistan" }
        "2691" { $ret = "Comoros" }
        "5399" { $ret = "Guantanamo Bay" }
        "6101" { $ret = "Cocos (Keeling) Islands" }
        default { $ret = $intCountryCode }
    }
    return $ret
}

#region Win32_Printer Functions
Function Get-Win32_Printer-Availability($uint16_Availability)
{
      switch ($uint16_Availability)
      {
            1 { $ret = "Other" }
            2 { $ret = "Unknown" }
            3 { $ret = "Running or Full Power" }
            4 { $ret = "Warning" }
            5 { $ret = "In Test" }
            6 { $ret = "Not Applicable" }
            7 { $ret = "Power Off" }
            8 { $ret = "Off Line" }
            9 { $ret = "Off Duty" }
            10 { $ret = "Degraded" }
            11 { $ret = "Not Installed" }
            12 { $ret = "Install Error" }
            13 { $ret = "Power Save - Unknown" }
            14 { $ret = "Power Save - Low Power Mode" }
            15 { $ret = "Power Save - Standby" }
            16 { $ret = "Power Cycle" }
            17 { $ret = "Power Save - Warning" }
            default { $ret = $uint16_Availability }
      }
      return $ret
}

Function Get-Win32_Printer-Capabilities($uint16_array_Capabilities)
{
      switch ($uint16_array_Capabilities)
      {
            0 { $ret = "Unknown" }
            1 { $ret = "Other" }
            2 { $ret = "Color Printing" }
            3 { $ret = "Duplex Printing" }
            4 { $ret = "Copies" }
            5 { $ret = "Collation" }
            6 { $ret = "Stapling" }
            7 { $ret = "Transparency Printing" }
            8 { $ret = "Punch" }
            9 { $ret = "Cover" }
            10 { $ret = "Bind" }
            11 { $ret = "Black and White Printing" }
            12 { $ret = "One-Sided" }
            13 { $ret = "Two-Sided Long Edge" }
            14 { $ret = "Two-Sided Short Edge" }
            15 { $ret = "Portrait" }
            16 { $ret = "Landscape" }
            17 { $ret = "Reverse Portrait" }
            18 { $ret = "Reverse Landscape" }
            19 { $ret = "Quality High" }
            20 { $ret = "Quality Normal" }
            21 { $ret = "Quality Low" }
            default { $ret = $uint16_array_Capabilities }
      }
      return $ret
}

Function Get-Win32_Printer-ConfigManagerErrorCode($uint32_ConfigManagerErrorCode)
{
      switch ($uint32_ConfigManagerErrorCode)
      {
            0 { $ret = "Device is working properly." }
            1 { $ret = "Device is not configured correctly." }
            2 { $ret = "Windows cannot load the driver for this device." }
            3 { $ret = "Driver for this device might be corrupted, or the  system may be  low on memory or other resources." }
            4 { $ret = "Device is not working properly. One of its drivers or the  registry might be corrupted." }
            5 { $ret = "Driver for the device requires a resource that Windows cannot manage." }
            6 { $ret = "Boot configuration for the device conflicts with other devices." }
            7 { $ret = "Cannot filter." }
            8 { $ret = "Driver loader for the device is missing." }
            9 { $ret = "Device is not working properly.  The controlling firmware is incorrectly reporting the resources for the device." }
            10 { $ret = "Device cannot start." }
            11 { $ret = "Device failed." }
            12 { $ret = "Device cannot find enough free resources to use." }
            13 { $ret = "Windows cannot verify the device's resources." }
            14 { $ret = "Device cannot work properly until the computer is restarted." }
            15 { $ret = "Device is not working properly due to a possible re-enumeration problem." }
            16 { $ret = "Windows cannot identify all of the resources that the device uses." }
            17 { $ret = "Device is requesting  an unknown resource type." }
            18 { $ret = "Device drivers must be reinstalled." }
            19 { $ret = "Failure using the VxD loader." }
            20 { $ret = "Registry might be corrupted." }
            21 { $ret = "System failure. If changing the device driver is ineffective, see the hardware documentation. Windows is removing the device." }
            22 { $ret = "Device is disabled." }
            23 { $ret = "System failure. If changing the device driver is ineffective, see the hardware documentation." }
            24 { $ret = "Device is not present,  not working properly, or does not have all of its drivers installed." }
            25 { $ret = "Windows is still setting up the device." }
            26 { $ret = "Windows is still setting up the device." }
            27 { $ret = "Device does not have a valid log configuration." }
            28 { $ret = "Device drivers   are not installed." }
            29 { $ret = "Device is disabled.  The device firmware   did not provide  the required resources." }
            30 { $ret = "Device is using an IRQ resource that another device is using." }
            31 { $ret = "Device is not working properly.  Windows cannot load the  required device drivers." }
            default { $ret = $uint32_ConfigManagerErrorCode }
      }
      return $ret
}

Function Get-Win32_Printer-CurrentCapabilities($uint16_array_CurrentCapabilities)
{
      switch ($uint16_array_CurrentCapabilities)
      {
            0 { $ret = "Unknown" }
            1 { $ret = "Other" }
            2 { $ret = "Color Printing" }
            3 { $ret = "Duplex Printing" }
            4 { $ret = "Copies" }
            5 { $ret = "Collation" }
            6 { $ret = "Stapling" }
            7 { $ret = "Transparency Printing" }
            8 { $ret = "Punch" }
            9 { $ret = "Cover" }
            10 { $ret = "Bind" }
            11 { $ret = "Black and White Printing" }
            12 { $ret = "One-Sided" }
            13 { $ret = "Two-Sided Long Edge" }
            14 { $ret = "Two-Sided Short Edge" }
            15 { $ret = "Portrait" }
            16 { $ret = "Landscape" }
            17 { $ret = "Reverse Portrait" }
            18 { $ret = "Reverse Landscape" }
            19 { $ret = "Quality High" }
            20 { $ret = "Quality Normal" }
            21 { $ret = "Quality Low" }
            default { $ret = $uint16_array_CurrentCapabilities }
      }
      return $ret
}

Function Get-Win32_Printer-CurrentLanguage($uint16_CurrentLanguage)
{
      switch ($uint16_CurrentLanguage)
      {
            1 { $ret = "Other" }
            2 { $ret = "Unknown" }
            3 { $ret = "PCL" }
            4 { $ret = "HPGL" }
            5 { $ret = "PJL" }
            6 { $ret = "PS" }
            7 { $ret = "PSPrinter" }
            8 { $ret = "IPDS" }
            9 { $ret = "PPDS" }
            10 { $ret = "EscapeP" }
            11 { $ret = "Epson" }
            12 { $ret = "DDIF" }
            13 { $ret = "Interpress" }
            14 { $ret = "ISO6429" }
            15 { $ret = "LineData" }
            16 { $ret = "DODCA" }
            17 { $ret = "REGIS" }
            18 { $ret = "SCS" }
            19 { $ret = "SPDL" }
            20 { $ret = "TEK4014" }
            21 { $ret = "PDS" }
            22 { $ret = "IGP" }
            23 { $ret = "CodeV" }
            24 { $ret = "DSCDSE" }
            25 { $ret = "WPS" }
            26 { $ret = "LN03" }
            27 { $ret = "CCITT" }
            28 { $ret = "QUIC" }
            29 { $ret = "CPAP" }
            30 { $ret = "DecPPL" }
            31 { $ret = "SimpleText" }
            32 { $ret = "NPAP" }
            33 { $ret = "DOC" }
            34 { $ret = "imPress" }
            35 { $ret = "Pinwriter" }
           36 { $ret = "NPDL" }
            37 { $ret = "NEC201PL" }
            38 { $ret = "Automatic" }
            39 { $ret = "Pages" }
            40 { $ret = "LIPS" }
            41 { $ret = "TIFF" }
            42 { $ret = "Diagnostic" }
            43 { $ret = "CaPSL" }
            44 { $ret = "EXCL" }
            45 { $ret = "LCDS" }
            46 { $ret = "XES" }
            47 { $ret = "MIME" }
            48 { $ret = "XPS" }
            49 { $ret = "HPGL2" }
            50 { $ret = "PCLXL" }
            default { $ret = $uint16_CurrentLanguage }
      }
      return $ret
}

Function Get-Win32_Printer-DefaultCapabilities($uint16_array_DefaultCapabilities)
{
      switch ($uint16_array_DefaultCapabilities)
      {
            0 { $ret = "Unknown" }
            1 { $ret = "Other" }
            2 { $ret = "Color Printing" }
            3 { $ret = "Duplex Printing" }
            4 { $ret = "Copies" }
            5 { $ret = "Collation" }
            6 { $ret = "Stapling" }
            7 { $ret = "Transparency Printing" }
            8 { $ret = "Punch" }
            9 { $ret = "Cover" }
            10 { $ret = "Bind" }
            11 { $ret = "Black and White Printing" }
            12 { $ret = "One-Sided" }
            13 { $ret = "Two-Sided Long Edge" }
            14 { $ret = "Two-Sided Short Edge" }
            15 { $ret = "Portrait" }
            16 { $ret = "Landscape" }
            17 { $ret = "Reverse Portrait" }
            18 { $ret = "Reverse Landscape" }
            19 { $ret = "Quality High" }
            20 { $ret = "Quality Normal" }
            21 { $ret = "Quality Low" }
            default { $ret = $uint16_array_DefaultCapabilities }
      }
      return $ret
}

Function Get-Win32_Printer-DefaultLanguage($uint16_DefaultLanguage)
{
      switch ($uint16_DefaultLanguage)
      {
            1 { $ret = "Other" }
            2 { $ret = "Unknown" }
            3 { $ret = "PCL" }
            4 { $ret = "HPGL" }
            5 { $ret = "PJL" }
            6 { $ret = "PS" }
            7 { $ret = "PSPrinter" }
            8 { $ret = "IPDS" }
            9 { $ret = "PPDS" }
            10 { $ret = "EscapeP" }
            11 { $ret = "Epson" }
            12 { $ret = "DDIF" }
            13 { $ret = "Interpress" }
            14 { $ret = "ISO6429" }
            15 { $ret = "LineData" }
            16 { $ret = "DODCA" }
            17 { $ret = "REGIS" }
            18 { $ret = "SCS" }
            19 { $ret = "SPDL" }
            20 { $ret = "TEK4014" }
            21 { $ret = "PDS" }
            22 { $ret = "IGP" }
            23 { $ret = "CodeV" }
            24 { $ret = "DSCDSE" }
            25 { $ret = "WPS" }
            26 { $ret = "LN03" }
            27 { $ret = "CCITT" }
            28 { $ret = "QUIC" }
            29 { $ret = "CPAP" }
            30 { $ret = "DecPPL" }
            31 { $ret = "SimpleText" }
            32 { $ret = "NPAP" }
            33 { $ret = "DOC" }
            34 { $ret = "imPress" }
            35 { $ret = "Pinwriter" }
            36 { $ret = "NPDL" }
            37 { $ret = "NEC201PL" }
            38 { $ret = "Automatic" }
            39 { $ret = "Pages" }
            40 { $ret = "LIPS" }
            41 { $ret = "TIFF" }
            42 { $ret = "Diagnostic" }
            43 { $ret = "CaPSL" }
            44 { $ret = "EXCL" }
            45 { $ret = "LCDS" }
            46 { $ret = "XES" }
            47 { $ret = "MIME" }
            48 { $ret = "XPS" }
            49 { $ret = "HPGL2" }
            50 { $ret = "PCLXL" }
            default { $ret = $uint16_DefaultLanguage }
      }
      return $ret
}

Function Get-Win32_Printer-DetectedErrorState($uint16_DetectedErrorState)
{
      switch ($uint16_DetectedErrorState)
      {
            0 { $ret = "Unknown" }
            1 { $ret = "Other" }
            2 { $ret = "No Error" }
            3 { $ret = "Low Paper" }
            4 { $ret = "No Paper" }
            5 { $ret = "Low Toner" }
            6 { $ret = "No Toner" }
            7 { $ret = "Door Open" }
            8 { $ret = "Jammed" }
            9 { $ret = "Offline" }
            10 { $ret = "Service Requested" }
            11 { $ret = "Output Bin Full" }
            default { $ret = $uint16_DetectedErrorState }
      }
      return $ret
}

Function Get-Win32_Printer-ExtendedPrinterStatus($uint16_ExtendedPrinterStatus)
{
      switch ($uint16_ExtendedPrinterStatus)
      {
            1 { $ret = "Other" }
            2 { $ret = "Unknown" }
            3 { $ret = "Idle" }
            4 { $ret = "Printing" }
            5 { $ret = "Warming Up" }
            6 { $ret = "Stopped Printing" }
            7 { $ret = "Offline" }
            8 { $ret = "Paused" }
            9 { $ret = "Error" }
            10 { $ret = "Busy" }
            11 { $ret = "Not Available" }
            12 { $ret = "Waiting" }
            13 { $ret = "Processing" }
            14 { $ret = "Initialization" }
            15 { $ret = "Power Save" }
            16 { $ret = "Pending Deletion" }
            17 { $ret = "I/O Active" }
            18 { $ret = "Manual Feed" }
            default { $ret = $uint16_ExtendedPrinterStatus }
      }
      return $ret
}

Function Get-Win32_Printer-LanguagesSupported($uint16_array_LanguagesSupported)
{
      switch ($uint16_array_LanguagesSupported)
      {
            1 { $ret = "Other" }
            2 { $ret = "Unknown" }
            3 { $ret = "PCL" }
            4 { $ret = "HPGL" }
            5 { $ret = "PJL" }
            6 { $ret = "PS" }
            7 { $ret = "PSPrinter" }
            8 { $ret = "IPDS" }
            9 { $ret = "PPDS" }
            10 { $ret = "EscapeP" }
            11 { $ret = "Epson" }
            12 { $ret = "DDIF" }
            13 { $ret = "Interpress" }
            14 { $ret = "ISO6429" }
            15 { $ret = "LineData" }
            16 { $ret = "DODCA" }
            17 { $ret = "REGIS" }
            18 { $ret = "SCS" }
            19 { $ret = "SPDL" }
            20 { $ret = "TEK4014" }
            21 { $ret = "PDS" }
            22 { $ret = "IGP" }
            23 { $ret = "CodeV" }
            24 { $ret = "DSCDSE" }
            25 { $ret = "WPS" }
            26 { $ret = "LN03" }
            27 { $ret = "CCITT" }
            28 { $ret = "QUIC" }
            29 { $ret = "CPAP" }
            30 { $ret = "DecPPL" }
            31 { $ret = "SimpleText" }
            32 { $ret = "NPAP" }
            33 { $ret = "DOC" }
            34 { $ret = "imPress" }
            35 { $ret = "Pinwriter" }
            36 { $ret = "NPDL" }
            37 { $ret = "NEC201PL" }
            38 { $ret = "Automatic" }
            39 { $ret = "Pages" }
            40 { $ret = "LIPS" }
            41 { $ret = "TIFF" }
            42 { $ret = "Diagnostic" }
            43 { $ret = "CaPSL" }
            44 { $ret = "EXCL" }
            45 { $ret = "LCDS" }
            46 { $ret = "XES" }
            47 { $ret = "MIME" }
            48 { $ret = "XPS" }
            49 { $ret = "HPGL2" }
            50 { $ret = "PCLXL" }
            default { $ret = $uint16_array_LanguagesSupported }
      }
      return $ret
}

Function Get-Win32_Printer-MarkingTechnology($uint16_MarkingTechnology)
{
      switch ($uint16_MarkingTechnology)
      {
            1 { $ret = "Other" }
            2 { $ret = "Unknown" }
           3 { $ret = "Electrophotographic LED" }
            4 { $ret = "Electrophotographic Laser" }
            5 { $ret = "Electrophotographic Other" }
            6 { $ret = "Impact Moving Head Dot Matrix 9pin" }
            7 { $ret = "Impact Moving Head Dot Matrix 24pin" }
            8 { $ret = "Impact Moving Head Dot Matrix Other" }
            9 { $ret = "Impact Moving Head Fully Formed" }
            10 { $ret = "Impact Band" }
            11 { $ret = "Impact Other" }
            12 { $ret = "Inkjet Aqueous" }
            13 { $ret = "Inkjet Solid" }
            14 { $ret = "Inkjet Other" }
            15 { $ret = "Pen" }
            16 { $ret = "Thermal Transfer" }
            17 { $ret = "Thermal Sensitive" }
            18 { $ret = "Thermal Diffusion" }
            19 { $ret = "Thermal Other" }
            20 { $ret = "Electroerosion" }
            21 { $ret = "Electrostatic" }
            22 { $ret = "Photographic Microfiche" }
            23 { $ret = "Photographic Imagesetter" }
            24 { $ret = "Photographic Other" }
            25 { $ret = "Ion Deposition" }
            26 { $ret = "eBeam" }
            27 { $ret = "Typesetter" }
            default { $ret = $uint16_MarkingTechnology }
      }
      return $ret
}

Function Get-Win32_Printer-PaperSizesSupported($uint16_array_PaperSizesSupported)
{
      switch ($uint16_array_PaperSizesSupported)
      {
            0 { $ret = "Unknown" }
            1 { $ret = "Other" }
            2 { $ret = "A" }
            3 { $ret = "B" }
            4 { $ret = "C" }
            5 { $ret = "D" }
            6 { $ret = "E" }
            7 { $ret = "Letter" }
            8 { $ret = "Legal" }
            9 { $ret = "NA-10x13-Envelope" }
            10 { $ret = "NA-9x12-Envelope" }
            11 { $ret = "NA-Number-10-Envelope" }
            12 { $ret = "NA-7x9-Envelope" }
            13 { $ret = "NA-9x11-Envelope" }
            14 { $ret = "NA-10x14-Envelope" }
            15 { $ret = "NA-Number-9-Envelope" }
            16 { $ret = "NA-6x9-Envelope" }
            17 { $ret = "NA-10x15-Envelope" }
            18 { $ret = "A0" }
            19 { $ret = "A1" }
            20 { $ret = "A2" }
            21 { $ret = "A3" }
            22 { $ret = "A4" }
            23 { $ret = "A5" }
            24 { $ret = "A6" }
            25 { $ret = "A7" }
            26 { $ret = "A8" }
            27 { $ret = "A9A10" }
            28 { $ret = "B0" }
            29 { $ret = "B1" }
            30 { $ret = "B2" }
            31 { $ret = "B3" }
            32 { $ret = "B4" }
            33 { $ret = "B5" }
            34 { $ret = "B6" }
            35 { $ret = "B7" }
            36 { $ret = "B8" }
            37 { $ret = "B9" }
            38 { $ret = "B10" }
            39 { $ret = "C0" }
            40 { $ret = "C1" }
            41 { $ret = "C2" }
            42 { $ret = "C3" }
            43 { $ret = "C4" }
            44 { $ret = "C5" }
            45 { $ret = "C6" }
            46 { $ret = "C7" }
            47 { $ret = "C8" }
            48 { $ret = "ISO-Designated" }
            49 { $ret = "JIS B0" }
            50 { $ret = "JIS B1" }
            51 { $ret = "JIS B2" }
            52 { $ret = "JIS B3" }
            53 { $ret = "JIS B4" }
            54 { $ret = "JIS B5" }
            55 { $ret = "JIS B6" }
            56 { $ret = "JIS B7" }
            57 { $ret = "JIS B8" }
            58 { $ret = "JIS B9" }
            59 { $ret = "JIS B10" }
            default { $ret = $uint16_array_PaperSizesSupported }
      }
      return $ret
}

Function Get-Win32_Printer-PowerManagementCapabilities($uint16_array_PowerManagementCapabilities)
{
      switch ($uint16_array_PowerManagementCapabilities)
      {
            0 { $ret = "Unknown" }
            1 { $ret = "Not Supported" }
            2 { $ret = "Disabled" }
            3 { $ret = "Enabled" }
            4 { $ret = "Power Saving Modes Entered Automatically" }
            5 { $ret = "Power State Settable" }
            6 { $ret = "Power Cycling Supported" }
            7 { $ret = "Timed Power-On Supported" }
            default { $ret = $uint16_array_PowerManagementCapabilities }
      }
      return $ret
}

Function Get-Win32_Printer-PrinterState($uint32_PrinterState)
{
      switch ($uint32_PrinterState)
      {
            1 { $ret = "Paused" }
            2 { $ret = "Error" }
            3 { $ret = "Pending Deletion" }
            4 { $ret = "Paper Jam" }
            5 { $ret = "Paper Out" }
            6 { $ret = "Manual Feed" }
            7 { $ret = "Paper Problem" }
            8 { $ret = "Offline" }
            9 { $ret = "I/O Active" }
            10 { $ret = "Busy" }
            11 { $ret = "Printing" }
            12 { $ret = "Output Bin Full" }
            13 { $ret = "Not Available" }
            14 { $ret = "Waiting" }
            15 { $ret = "Processing" }
            16 { $ret = "Initialization" }
            17 { $ret = "Warming Up" }
            18 { $ret = "Toner Low" }
            19 { $ret = "No Toner" }
            20 { $ret = "Page Punt" }
            21 { $ret = "User Intervention Required" }
            22 { $ret = "Out of Memory" }
            23 { $ret = "Door Open" }
            24 { $ret = "Server_Unknown" }
            25 { $ret = "Power Save" }
            default { $ret = $uint32_PrinterState }
      }
      return $ret
}

Function Get-Win32_Printer-PrinterStatus($uint16_PrinterStatus)
{
      switch ($uint16_PrinterStatus)
      {
            1 { $ret = "Other" }
            2 { $ret = "Unknown" }
            3 { $ret = "Idle" }
            4 { $ret = "Printing" }
            5 { $ret = "Warming Up" }
            6 { $ret = "Stopped printing" }
            7 { $ret = "Offline" }
            default { $ret = $uint16_PrinterStatus }
      }
      return $ret
}

Function Get-Win32_Printer-StatusInfo($uint16_StatusInfo)
{
      switch ($uint16_StatusInfo)
      {
            1 { $ret = "Other" }
            2 { $ret = "Unknown" }
            3 { $ret = "Enabled" }
            4 { $ret = "Disabled" }
            5 { $ret = "Not Applicable" }
            default { $ret = $uint16_StatusInfo }
      }
      return $ret
}
#endregion Win32_Printer Functions

function WMIAssociationGrouptoPart($Group,$GroupList)
{
	$ret = ""
	ForEach($G in $GroupList) {
		$GroupComponent = $G.GroupComponent.split("=")[1].replace('"',"")
		$PartComponent  = $G.PartComponent.split("=")[1].replace('"',"") 		
		If ($GroupComponent -eq $Group) {
			$ret += $PartComponent + "`t"
		}
	}
	return $ret.trim().replace("`t","<BR>")
}

function WMIAssociationPartToGroup($Part,$Group)
{
	$ret = ""
	ForEach($G in $Group) {
		$GroupComponent = $G.GroupComponent.split("=")[1].replace('"',"")
		$PartComponent  = $G.PartComponent.split("=")[1].replace('"',"") 
		If ($PartComponent -eq $Part) {
			$ret += $GroupComponent + "`t"
		}
	}
	return $ret.trim().replace("`t","<BR>")
}

function get_WMIDate($WMI_Date)
{
    $objDate = WMIDateStringToDateTime($WMI_Date) 
    return $objDate.ToLongDateString() + " " + $objDate.ToLongTimeString()
}

function YesNo($boolean)
{
    if ( $boolean ){   
        $ret = "Yes" 
    } else {
        $ret = "No" 
    }
    return $ret
}
function get_DomainOrWorkgroup($strDomain,$strWorkgroup)
{
      if ($strDomain -eq "") {
            $ret = $strWorkgroup }
      else {
            $ret = $strDomain[0] }
      
      return $ret
}
function WMIDateStringToDateTime( [String] $strWmiDate ){
    # credit to Grey Lyon - http://gallery.technet.microsoft.com/scriptcenter/2c93d198-ec69-4c04-958b-bc089eeaa0d4
    $strWmiDate = $strWmiDate.Trim()
    $iYear   = [Int32]::Parse($strWmiDate.SubString( 0, 4))
    $iMonth  = [Int32]::Parse($strWmiDate.SubString( 4, 2))
    $iDay    = [Int32]::Parse($strWmiDate.SubString( 6, 2))
    if ($strWmiDate.length -gt 8) {
        $iHour   = [Int32]::Parse($strWmiDate.SubString( 8, 2))
        $iMinute = [Int32]::Parse($strWmiDate.SubString(10, 2))
        $iSecond = [Int32]::Parse($strWmiDate.SubString(12, 2))
        # decimal point is at $strWmiDate.Substring(14, 1)
        $iMicroseconds = [Int32]::Parse($strWmiDate.Substring(15, 6))
        $iMilliseconds = $iMicroseconds / 1000
        $iUtcOffsetMinutes = [Int32]::Parse($strWmiDate.Substring(21, 4))
    } else {
        $iHour   = 0
        $iMinute = 0
        $iSecond = 0
        $iMicroseconds = 0
        $iMilliseconds = 0
        $iUtcOffsetMinutes = 0
    }
    if ( $iUtcOffsetMinutes -ne 0 )
    {
        $dtkind = [DateTimeKind]::Local
    } else {
        $dtkind = [DateTimeKind]::Utc
    }
    return (New-Object -TypeName DateTime `
                -ArgumentList $iYear, $iMonth, $iDay, $iHour, $iMinute, $iSecond, $iMilliseconds, $dtkind)
}
function WMIShortDateStringToDate( [String] $strWmiDate )
{
    # credit to Grey Lyon - http://gallery.technet.microsoft.com/scriptcenter/2c93d198-ec69-4c04-958b-bc089eeaa0d4
    $strWmiDate = $strWmiDate.Trim()
    $iYear   = [Int32]::Parse($strWmiDate.SubString( 0, 4))
    $iMonth  = [Int32]::Parse($strWmiDate.SubString( 4, 2))
    $iDay    = [Int32]::Parse($strWmiDate.SubString( 6, 2))
    $iHour   = 0
    $iMinute = 0
    $iSecond = 0
    $iMicroseconds = 0
    $iMilliseconds = 0
    $iUtcOffsetMinutes = 0

    $dtkind = [DateTimeKind]::Utc

    $date = New-Object -TypeName DateTime `
                      -ArgumentList $iYear, $iMonth, $iDay, `
                                    $iHour, $iMinute, $iSecond, `
                                    $iMilliseconds, $dtkind
    return (Get-Date $date).ToShortDateString()                                               
}

function get_date([string]$strString)
{     
      $ret = $strString.split(" ")[0]
      return $ret
}


Function Capitise($InString) {
	$ret = ""
	$CharArray = $InString.tocharArray()
	For($i=0; $i -le $InString.length; $i++) {
		$Char = $CharArray[$i]
		$CharNext = If ($i -ne $InString.length-1) { $CharArray[$i+1] } else {$Char}
		$CharLast = If ($i -gt 0) { $CharArray[$i-1] } else {$Char}		
		
		If (65..90 -contains [byte]$Char) { # if upper case
			If ((65..90 -contains [byte]$CharLast) -and (65..90 -contains [byte]$CharNext)) { 
				$ret += $Char
			} else {
				$ret += " " + $Char
			}
		} else {			
			$ret += $Char
		}
	}
	return $ret.trim()
}

Function RenameHeaders( $PSObject ) {
    If ($PSObject -ne $null) {
        $ret = @()

        Foreach ($line in $PSObject) {
            $NewObject = New-Object –TypeName PSObject
            foreach ($header in $line.psobject.properties) {
                $NewObject | Add-Member Noteproperty –Name $(Capitise $header.name) –Value $header.value
            }
            $ret += $NewObject 
        }
        $ret
    }
}

Function ReportProgress($OutputText)
{
      Write-Host $OutputText
}

Function OutputHTML($HeaderType,$Header,$TableType,$Win32) {      
	Write-progress -Activity ("Collecting Information : " + $Computer) -status "Writing $Header" -PercentComplete ($percent * $global:i++)

	$Object = (Get-Variable $win32 -ErrorAction "SilentlyContinue" ).Value
	if ($Object -ne $null) {
	    Switch ($HeaderType) {
			1 { $ret += "<H1><a name=""$Win32""></a>$Header</H1><BR>" }
			2 { $ret += "<H2><a name=""$Win32"">$Header</H2><BR>" }
			3 { $ret += "<H3><a name=""$Win32"">$Header</H3><BR>" }
			4 { $ret += "<H4><a name=""$Win32"">$Header</H4><BR>" }
	    }
	    $ret += ConvertTo-HTMLTable $TableType $Object 
	    $ret
	}
}

function GenerateRowHTML($Content) {
      $return = [string]""
      $Content.psobject.properties | foreach {                          
            $return += "<tr>"
            if ($headers) {
                  $class = ">"
            } else {
                  $class = " class=""alt"">" 
            }
            $return += "<td width=""30%""" + $class + ($_.name) + "</td>"
            if ($_.Value -ne $null) {
                  $Value = $_.Value

                  Switch ($Value.gettype().tostring()) {
                        "System.Boolean" {
                              If ($Value) {
                                    $return += "<td>Yes</td>" 
                              } else {
                                    $return += "<td>No</td>" 
                              } }
                        "System.String" {
                              $return += "<td>" + ($Value) + "</td>" }
                        "System.Int16" {
                              $return += "<td>" + ("{0:N0}" -f $Value) + "</td>" }
                        "System.UInt16" {
                              $return += "<td>" + ("{0:N0}" -f $Value) + "</td>" }
                        "System.Int32" {
                              $return += "<td>" + ("{0:N0}" -f $Value) + "</td>" }
                        "System.UInt32" {
                              $return += "<td>" + ("{0:N0}" -f $Value) + "</td>" }
                        "System.Object[]" {
                              $return += "<td><table id=""table1"">"  + $vbcrlf
                              $Headers = $Value | Get-Member -type NoteProperty
                              $return += "<tr>"
                              $Headers | foreach {
                                    $return += "<td>" + $_.name + "</td>"
                              }
                              $return += "</tr>" + $vbcrlf                          
                              
                              Foreach($record in $Value) {
                                    $return += "<tr>"
                                    $Headers | foreach {
                                          $return += "<td>" + ($record.($_.name)) + "</td>"
                                    }
                                    $return += "</tr>" + $vbcrlf                          
                              } 
                              $return += "</table></td>"  + $vbcrlf
                              }
                        default {
                              #Write-Host $Value $Value.gettype().tostring()
                              $return += "<td>" + ($Value) + "</td>" }
                  }
            }           
            $return += "</tr>" + $vbcrlf                          
      }
      return $return
}

function ConvertTo-HTMLTable($Headers,$tblContents) {
      $alt = $true
      $html = "<table id=""table1"">"  + $vbcrlf

      if($tblContents -ne $null ) {
            if ($headers) {   
                  if ($tblContents.gettype().BaseType.Name -eq "Array") 
                  {
                        $Header = $tblContents.SyncRoot[0]
                  } else {
                        $Header = $tblContents
                  }
                  $tblHeaders = $Header.psobject.properties.name | Select-Object @{n='Name';e={$_}}
      
                  $html += "<tr>"
                  foreach ($tblHeader in $tblHeaders)
                  {
                        $html += "<th>" + $tblHeader.name + "</th>"
                  }
                  $html += "</tr>" + $vbcrlf    
                  
                  foreach($Line in $tblContents)
                  {
                        if ($alt) {
                              $html += "<tr>"
                              $alt = $False
                        } else {
                              $html += "<tr class=""alt"">"
                              $alt = $True
                        }
                        
                        foreach ($tblHeader in $tblHeaders)
                        {
                              $html += "<td>" + ($Line.($tblHeader.name) -replace $vbCrLf,"<BR>") + "</td>"
                        }
                        $html += "</tr>" + $vbcrlf                            
                  }
            } else {
                  If ($tblContents.gettype().tostring() -eq "System.Management.Automation.PSCustomObject" ) {
                        $html += GenerateRowHTML $tblContents
                  }
                  If ($tblContents.gettype().tostring() -eq "System.Object[]" ) {
                        Foreach ($tblContent in $tblContents) {
                              $html += GenerateRowHTML $tblContent
                              $html += "<tr><td colspan =2><BR></td></tr>" + $vbcrlf                              
                        }
                  }
            }
      }
      $html += "</table><BR>"  + $vbcrlf
      $html
}

Function MyGet-WmiObject([string]$computername,[string]$namespace="root\cimv2",[string]$class,[int]$timeout=60,[string]$username="",[string]$password="") { 
      $ConnectOptions = new-object System.Management.ConnectionOptions 
      $EnuOptions = new-object System.Management.EnumerationOptions

      $WMItimeout = (new-timespan -seconds $timeout)
      $EnuOptions.set_timeout($WMItimeout)

      if ($Username -ne "") {
            if ($username.contains("\")) {
                  $domain = $username.split("\")[0]
                  $username = $username.split("\")[1]
            } else {
                  $domain = $computername
            }
            $ConnectOptions.Username = $username
            $ConnectOptions.Password = $password
            $ConnectOptions.Authority = ("ntlmdomain:" + $domain)
      }

      $Scope = new-object System.Management.ManagementScope("\\" + $computername + "\" + $namespace ), $ConnectOptions 
      $ErrorActionPreference = "SilentlyContinue"
      $ret = $Scope.Connect() | Out-Null

      $query = new-object System.Management.ObjectQuery $querystring 
      $search = new-object System.Management.ManagementObjectSearcher 
      $search.set_options($EnuOptions) 
      $search.Scope = $Scope
      $search.Query = "SELECT * FROM " + $class 
      trap { 
            exit
      } 
      $result = $search.get()

      return $result 
}
#EndRegion Functions

$AttributesToRemove = @( "__GENUS", "__CLASS", "__SUPERCLASS", "__DYNASTY",  "__RELPATH", "__PROPERTY_COUNT", "__DERIVATION", "__SERVER", 
                                                "__NAMESPACE", "__PATH", "Scope", "Path", "Options", "ClassPath", "Properties", "SystemProperties", "Qualifiers",
                                                "Site", "Container")
If ($NoPing) {
      $Ping = $True
} else {
      $Ping = (Test-Connection -count 1 -ComputerName $Computer -ea silentlycontinue).StatusCode -eq 0
}
if ($Ping) {
    If (!(Test-Path $path)) {
        Write-Host "Output path $path does not exist"
        exit
    }
    If ($Credential -ne $null) {
        $WMICmd = " -Credential `$Credential"     
        $Username = $Credential.GetNetworkCredential().username
        $Password = $Credential.GetNetworkCredential().password
    }
    if ($Username -ne "" -and $cred -eq $null) {
        if ($Password -ne "") {
              $cred = New-Object -typename System.Management.Automation.PSCredential -argumentlist @($username, (ConvertTo-SecureString -String $password -AsPlainText -Force))
        }
        else {
              $Cred = Get-Credential $Username
        }
        $WMICmd = " -Credential `$Cred "
    }

    $Win32_Objects = MyGet-WmiObject -ComputerName $Computer -Class "meta_class" -Username $Username -Password $Password
    $Namespaces    = MyGet-WmiObject -ComputerName $Computer  -namespace "root" -class "__Namespace" -Username $Username -Password $Password | select name

    if ($Win32_Objects -ne $null ) {          
        $Win32_Objects = $Win32_Objects | select name | Where-Object {$_.name -like "win32_*"} | select -ExpandProperty Name
        $Cmd = "`$Reg = Get-WmiObject -list  -ComputerName $Computer -namespace root\default $WMICMD "
        Invoke-Expression $Cmd 
        $Reg = $Reg | Where-Object { $_.name -eq "StdRegProv" }
        
        $Win32_Report = "Win32_ComputerSystem", "Win32_OperatingSystem"
        Foreach($value in ($Win32Extra+$Win32Hardware+$Win32Software+$Win32Storage+$Win32Network+$Win32User+$Win32Misc)) {
            If ($value.Contains("#")) {
                $value = $value.Split("#")[0]
            }
            Remove-Variable $value -ErrorAction "SilentlyContinue"                
        }
        
        If ($ReportExtra)    { $Win32_Report += $Win32Extra }
        if ($ReportHardware) { $Win32_Report += $Win32Hardware }
        if ($ReportSoftware) { $Win32_Report += $Win32Software }             
        if ($ReportStorage)  { $Win32_Report += $Win32Storage }             
        if ($ReportNetwork)  { $Win32_Report += $Win32Network }
        if ($ReportUser)     { $Win32_Report += $Win32User }
        if ($ReportMisc)     { $Win32_Report += $Win32Misc }
              
        $percent = (100/$Win32_Report.count)
        # Get Required WMI Objects default namespace
        foreach ($Win32 in $Win32_Report) {
            $filter = ""
            $NameSpace = ""
            If ($Win32.Contains("#")) {
                  $Info = $Win32.Split("#")
                  $Win32 = $Info[0]
                  $Filter = $Info[1] 
                  $NameSpace = $Info[2]
                  if ($Filter -ne "") {
                        $Filter = "-filter " + $Filter
                  }
                  if (($NameSpace -ne "") -and ($Namespaces.name -contains $NameSpace.Replace("root\",""))) {
                        $NameSpace = "-namespace " + $NameSpace
                  } else {
                        $NameSpace = ""
                  }                             
                }
            if (($Win32_Objects -contains $Win32) -or ($NameSpace -ne "")) {
                Write-progress -Activity ("Collecting Information : " + $Computer) -status "Reading $Win32" -PercentComplete ($percent * $i++)
                $Cmd = "`$$Win32 = Get-WmiObject -ComputerName $Computer -class $Win32 $WMICmd $Filter $NameSpace -erroraction 'silentlycontinue'"  
                Invoke-Expression $Cmd 
            }
        }
        Write-progress -Activity ("Collecting Information : " + $Computer) -status "Reading additional information" -PercentComplete 100
        $key = $reg.GetStringValue($HKLM,"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon","DefaultDomainName")
        $LastUserDomain = $key.sValue
        $key = $reg.GetStringValue($HKLM,"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon","DefaultUserName")
        $LastUser = $key.sValue
        $key = $reg.GetStringValue($HKLM,"SYSTEM\CurrentControlSet\Control\Print\Printers","DefaultSpoolDirectory")
        $PrintSpoolerLocation = $key.sValue
        
        if ($LastUserDomain -ne $null) {
              $LastUser = $LastUserDomain + "\" + $LastUser
        }

        if ($Win32_ComputerSystem."Part Of Domain" -eq "Yes") {
              $DomainType = "domain" 
        } else {
              $DomainType = "workgroup" 
        }

        $i = 0
        # Tidy up objects
        $Win32_ComputerSystem   = $Win32_ComputerSystem |  
                Select-Object -Property Name,Caption,Manufacturer,Model,                  
                        @{n='Part Of Domain';e={YesNo($_.PartOfDomain)}}, 
                        @{n='Domain or Workgroup';e={get_DomainOrWorkgroup($_.Domain,$_.Workgroup)}},
                        @{n='Architecture';e={$_.SystemType}}, 
                        @{n='Domain Role';e={get_Win32_ComputerSystem_DomainRole($_.DomainRole)}}, 
                        @{n='Page File Automatic';e={YesNo($_.AutomaticManagedPagefile)}},                     
                        @{n='Chassis Boot State';e={get_Win32_ComputerSystem_ChassisBootupState($_.ChassisBootupState)}}, 
                        @{n='Daylight Saving';e={YesNo($_.DaylightInEffect)}}, 
                        @{n='Power State';e={get_Win32_ComputerSystem_PowerState($_.PowerState)}}, 
                        @{n='Power Supply State';e={get_Win32_ComputerSystem_PowerSupplyState($_.PowerSupplyState)}}, 
                        @{n='Memory';e={"{0:N0}" -f ($_.TotalPhysicalMemory / 1MB -as [int]) + " MB"}}, 
                        @{n='Processors';e={$_.NumberOfProcessors}},
                        @{n='Cores';e={$_.NumberOfLogicalProcessors}},
                        @{n='PC System Type';e={get_Win32_ComputerSystem_PCSystemType($_.PCSystemType)}},Status,
                        @{n='Primary Owner Name';e={$_.PrimaryOwnerName}},
                        @{n='Primary Owner Contact';e={$_.PrimaryOwnerContact}},Roles,
                        @{n='Bootup State';e={$_.BootupState}}
        $Win32_OperatingSystem = $Win32_OperatingSystem | 
                Select-Object -Property @{n='OS';e={$_.Caption}},@{n='Service Pack';e={$_.CSDVersion}},Manufacturer,
                        @{n='Registered User';e={($_.RegisteredUser)}},Organization,
                        @{n='Architecture';e={($_.OSArchitecture)}},
                        @{n='Operating System SKU';e={get_Win32_OperatingSystem_OperatingSystemSKU($_.OperatingSystemSKU)}},
                        @{n='Install Date';e={Get_WMIDate($_.InstallDate)}},
                        @{n='Last Boot';e={Get_WMIDate($_.LastBootUpTime)}},
                        @{n='Server Time';e={Get_WMIDate($_.LocalDateTime)}},
                        @{n='Windows Directory';e={($_.WindowsDirectory)}},
                        @{n='OS Language';e={get_Win32_OperatingSystem_OSLanguage($_.OSLanguage )}},
                        @{n='PAE Enabled';e={YesNo($_.PAEEnabled)}},                                        
                        @{n='Build Type';e={($_.BuildType)}},
                        @{n='Version';e={($_.Version)}},
                        @{n='Boot Device';e={($_.BootDevice)}},
                        @{n='Country Code';e={get_Win32_OperatingSystem_CountryCode($_.CountryCode)}},
                        @{n='Code Set';e={($_.CodeSet)}},Status,Locale,
                        @{n='Foreground Application Boost ';e={get_Win32_OperatingSystem_ForegroundApplicationBoost($_.ForegroundApplicationBoost )}}
        if ($ReportHardware) {
            $Win32_Processor  = $Win32_Processor | 
                    Select-Object -Property "Manufacturer","Description","Name",
                        @{n='Number of Cores';e={($_.NumberOfCores)}},                                          
                        @{n='Number of Logical Processors';e={($_.NumberOfLogicalProcessors)}},
                        @{n='Maximum speed of the processor';e={"{0:N0}" -f ($_.MaxClockSpeed)}},
                        @{n='Availability';e={Get-Win32_Processor-Availability($_.Availability)}},L2CacheSize, ExtClock
            $Win32_BIOS  = $Win32_BIOS | 
                    Select-Object -Property "Manufacturer","Name",
                        @{n='Serial Number';e={($_.SerialNumber)}},"Version","Status",
                        @{n='Current Language';e={($_.CurrentLanguage)}},
                        @{n='Characteristics';e={([system.String]::join(", ",$_.BiosCharacteristics))}},
                        @{n='SM BIOS Version';e={($_.SMBIOSBIOSVersion)}},                                         
                        @{n='Major Version';e={($_.SMBIOSMajorVersion)}},
                        @{n='Minor Version';e={($_.SMBIOSMinorVersion)}}
            $Win32_BaseBoard = $Win32_BaseBoard |
                    Select-Object -Property Name,Manufacturer,Product,Version,
                        @{n='Serial Number';e={($_.SerialNumber)}}
            $Win32_ComputerSystemProduct = $Win32_ComputerSystemProduct |
                    Select-Object -Property Name,Vendor,
                        @{n='Identifying Number';e={($_.IdentifyingNumber)}}
            $Win32_SystemEnclosure = $Win32_SystemEnclosure |
                    Select-Object -Property @{n='Chassis Types';e={Get-Win32_SystemEnclosure-ChassisTypes($_.ChassisTypes)}},
                        @{n='Serial Number';e={($_.SerialNumber)}},Manufacturer,Version,SMBIOSAssetTag
            $Win32_PhysicalMemory = $Win32_PhysicalMemory |
                    Select-Object -Property @{n='Bank';e={$_.BankLabel}},@{n='Capacity';e={"{0:N0}" -f ($_.Capacity / 1MB -as [int]) + " MB"}},SerialNumber,
                        @{n='Form Factor';e={Get-Win32_PhysicalMemory-FormFactor($_.FormFactor)}},
                        @{n='Memory Type';e={Get-Win32_PhysicalMemory-MemoryType($_.MemoryType)}}
            $Win32_CDROMDrive = $Win32_CDROMDrive |
                    Select-Object -Property Name,Drive,Manufacturer
            $Win32_SoundDevice = $Win32_SoundDevice | 
                    Select-Object -Property Name, Manufacturer                        
            $Win32_VideoController = $Win32_VideoController |
                    Select-Object -Property @{n='Compatibility';e={$_.AdapterCompatibility}}, @{n='RAM';e={"{0:N0}" -f ($_.AdapterRAM / 1MB -as [int]) + " MB"}}, Name
            $Win32_TapeDrive = $Win32_TapeDrive | 
                    Select-Object -Property Name, Description, Manufacturer            
            $Win32_Keyboard = $Win32_Keyboard |
                    Select-Object -Property Name,Layout
            $Win32_PointingDevice = $Win32_PointingDevice |
                    Select-Object -Property Name,Manufacturer,Status                                
        }             
        if ($ReportSoftware) {
            $Win32_Product = $Win32_Product | 
                    Select-Object -Property Caption,Vendor,Version,@{l="Install Date";e={WMIShortDateStringToDate($_.InstallDate)}} | Sort-Object Caption                                 
            $Win32_OptionalFeature = $Win32_OptionalFeature |
                    Where-Object {$_.InstallState -eq 1} |
                        Select-Object -Property Name | Sort-Object Name
            $Win32_QuickFixEngineering = $Win32_QuickFixEngineering | where { $_.HotfixID -ne "File 1" } | Sort-Object HotFixID |
                    Select-Object -Property @{n='Patch ID';e={$_.HotFixID}},Description,@{l="Install Date";e={([DateTime]$_.psbase.properties["installedon"].value).tostring("d")}} 
        }             
        if ($ReportStorage) {
            $Win32_DiskDrive = $Win32_DiskDrive | select * -ExcludeProperty $AttributesToRemove
            $Win32_DiskDriveToDiskPartition = $Win32_DiskDriveToDiskPartition | select * -ExcludeProperty $AttributesToRemove
            $Win32_DiskPartition = $Win32_DiskPartition | select * -ExcludeProperty $AttributesToRemove
            $Win32_LogicalDisk = $Win32_LogicalDisk | select * -ExcludeProperty $AttributesToRemove
            $Win32_LogicalDiskToPartition = $Win32_LogicalDiskToPartition | select * -ExcludeProperty $AttributesToRemove
        
            $Disks = @()
            $Win32_DiskDrive = $Win32_DiskDrive | Sort-Object "DeviceID" 
            foreach($Win32_disk in $Win32_DiskDrive)
            {                       
                $Partitions = @()             
                $MyDisk = New-Object –TypeName PSObject
                $MyDisk | Add-Member Noteproperty –Name Caption –Value $Win32_disk.Caption
                $MyDisk | Add-Member Noteproperty –Name DeviceID –Value $Win32_disk.DeviceID
                $MyDisk | Add-Member Noteproperty –Name InterfaceType –Value $Win32_disk.InterfaceType
                $MyDisk | Add-Member Noteproperty –Name Size –Value ("{0:N0}" -f ($Win32_disk.Size / 1GB -as [int]) + " GB")
                
                foreach($DiskDriveToDiskPartition in ($Win32_DiskDriveToDiskPartition | where {$_.Antecedent.contains($Win32_disk.DeviceID.Replace("\","\\")) -eq $true}))
                {
                    if ($DiskDriveToDiskPartition -ne $null)
                    {
                        $PartName = $DiskDriveToDiskPartition.Dependent.split("=")[1] -replace($Quotes,"")
                        foreach($LogicalDiskToPartition in ($Win32_LogicalDiskToPartition | where {$_.Antecedent.contains($PartName) -eq $true}))
                        {                                               
                            $LogicalDisk = $Win32_LogicalDisk | where {$_.Caption -eq ($LogicalDiskToPartition.Dependent.split("=")[1] -replace($Quotes,""))}
                            $MyPartition = New-Object –TypeName PSObject
                            $MyPartition | Add-Member Noteproperty –Name Drive –Value $LogicalDisk.DeviceID     
                            $MyPartition | Add-Member Noteproperty –Name Volume –Value $LogicalDisk.VolumeName
                            $MyPartition | Add-Member Noteproperty –Name Size –Value ("{0:N0}" -f ($LogicalDisk.Size / 1GB -as [int]) + " GB")
                            $MyPartition | Add-Member Noteproperty –Name "Free Space" –Value ("{0:N0}" -f ($LogicalDisk.FreeSpace / 1GB -as [int]) + " GB")
                            $MyPartition | Add-Member Noteproperty –Name "File System" –Value $LogicalDisk.FileSystem                            
                            $Partitions += $MyPartition 
                        }                                         
                    }                                         
                }           
                $MyDisk | Add-Member Noteproperty –Name Partitions –Value $Partitions
                $Disks += $MyDisk
            }                             
        }             
        if ($ReportNetwork) {
            $Win32_NetworkAdapterConfiguration  = $Win32_NetworkAdapterConfiguration | where {$_.IPEnabled -eq "TRUE"} |
                Select-Object -Property @{n='IP Addresses';e={[string]::join(";",$_.IPAddress)}},Description,DNSDomain,
                    @{n='IP Subnet';e={[string]::join(";",$_.IPSubnet)}},
                    @{n='Gateway';e={[string]::join(";",$_.DefaultIPGateway)}},
                    @{n='DNS Servers';e={[string]::join(";",$_.DNSServerSearchOrder)}},
                    @{n='DNS Suffix';e={[string]::join(";",$_.DNSDomainSuffixSearchOrder)}},
                    @{n='MAC Address';e={[string]::join(";",$_.MACAddress)}},
                    @{n='WINS';e={[string]::join(";",@($_.WINSPrimaryServer,$_.WINSSecondaryServer))}},
                    @{n='DHCP Enabled';e={YesNo($_.DHCPEnabled)}},
                    @{n='DHCP Server';e={$_.DHCPServer}},
                    @{n='DHCP Lease Obtained';e={$_.DHCPLeaseObtained}},
                    @{n='DHCP Lease Expires';e={$_.DHCPLeaseExpires}},
                    @{n='Register in DNS';e={YesNo($_.DomainDNSRegistrationEnabled)}}
            $Win32_NetworkAdapter = $Win32_NetworkAdapter | where{$_.PhysicalAdapter -eq "True"} |  
                Select-Object NetConnectionID,Name,AdapterType,MACAddress,                                                   
                    @{n='Maximum Speed';e={("{0:N0}" -f ($_.Speed / 1MB -as [int]) + " MB")}},
                    @{n='Connection Status';e={get_Win32_NetworkAdapter_NetConnectionStatus($_.NetConnectionStatus)}},
                    @{n='Availability';e={get_Win32_NetworkAdapter_Availability($_.Availability)}}
            $Win32_IP4RouteTable = $Win32_IP4RouteTable |
                Select-Object -Property Destination, Mask, NextHop
        }
        if ($ReportMisc) {
            $Win32_NTEventLogFile = $Win32_NTEventLogFile |
                Select-Object -Property @{n='Log File Name';e={$_.LogFileName}},Name, 
                    @{n='Overwrite Policy';e={$_.OverwritePolicy}},
                    @{n='Maximum File Size';e={"{0:N0}" -f ($_.'MaxFileSize'/ 1MB -as [int]) + " MB"}}
            $Win32_Printer = $Win32_Printer |
                Select-Object -Property Name, @{n='Driver Name';e={$_.DriverName}}, @{n='Port Name';e={$_.Portname}},
                    @{n='Published';e={YesNo($_.Published)}}
            $Win32_TimeZone = $Win32_TimeZone | 
                Select-Object -Property @{n='Time Zone';e={$_.Description}},@{n='Daylight Name';e={($_.DaylightName)}}
            $Win32_Process = $Win32_Process |
                Select-Object -Property Caption, ExecutablePath
            $Win32_Service = $Win32_Service | Where-Object {$_.ServiceType -eq 'Share Process' -or $_.ServiceType -eq 'Own Process'}  |
                Select-Object -Property Caption, Started, @{n='Startup Mode';e={$_.StartMode}}, @{n='Start Name';e={$_.StartName}}                                          
            $Win32_Share = $Win32_Share |
                 Select-Object -Property Name, Description, Path
            $Win32_StartupCommand = $Win32_StartupCommand |
                Select-Object -Property Command, Name, User
            $Win32_PageFile = $Win32_PageFile | 
                Select-Object -Property "Caption",
                    @{n='Size';e={"{0:N0}" -f ($_.FileSize / 1MB -as [int]) + " MB"}},
                    @{n='Maximum Size';e={"{0:N0}" -f ($_.MaximumSize) + " MB"}},"Status"
            $Win32_Registry = $Win32_Registry |
                Select-Object -Property @{n='Current Size';e={"{0:N0}" -f ($_.CurrentSize) + " MB"}}, 
                    @{n='Maximum Size';e={"{0:N0}" -f ($_.MaximumSize) + " MB"}},
                    @{n='Proposed Size';e={"{0:N0}" -f ($_.ProposedSize) + " MB"}},"Status"
            $Win32_Environment = $Win32_Environment | where { $_.username -like "*system*" } |
                Select-Object -Property "Name",@{n='System Variable';e={YesNo($_.SystemVariable)}},Description,@{n='Variable Value';e={$_.VariableValue}}
            $Win32_NTDomain = RenameHeaders ($Win32_NTDomain | 
                Select-Object -Property "DomainName","ClientSiteName","DcSiteName","DnsForestName","DomainControllerAddress","DomainControllerAddressType")
            $Win32_TCPIPPrinterPort = $Win32_TCPIPPrinterPort | 
                Select-Object -Property Name,HostAddress,Description,@{n='Port Number';e={$_.PortNumber}}
            $Win32_PrinterDriver = $Win32_PrinterDriver |
                Select-Object -Property Name,MonitorName,Description,SupportedPlatform,Version
                                      
            $key = $reg.GetStringValue($HKLM,"SYSTEM\CurrentControlSet\Control\Print\Printers","DefaultSpoolDirectory")
            $PrintSpoolerLocation = $key.sValue
              
            $Win32_General =  [PSCustomObject]@{
                "Printer Spool Location" = $PrintSpoolerLocation;
                "Last Logon" = $LastUserDomain + "\" + $LastUser
            }                       
        }
        if ($ReportUser) {
              $Win32_Group = $Win32_Group |
                                Select-Object -Property Name, Description
              $Win32_UserAccount = $Win32_UserAccount |
                                Select-Object -Property Name, Description
        }
        if ($ReportExtra) {
            If ($MicrosoftNLB_Cluster -ne $null) {
                $MicrosoftNLB_Cluster = RenameHeaders ($MicrosoftNLB_Cluster |
                                                        Select-Object -Property Name,InterconnectAddress,ClusterState)
                $MicrosoftNLB_ClusterSetting = RenameHeaders ($MicrosoftNLB_ClusterSetting |
                                                        Select-Object -Property ClusterIPAddress, ClusterNetworkMask, ClusterMACAddres,MulticastSupportEnabled)
                $MicrosoftNLB_Node = RenameHeaders ($MicrosoftNLB_Node |
                                                        Select-Object -Property Name,HostPriority,DedicatedIPAddress,StatusCode)
                $MicrosoftNLB_NodeSetting = RenameHeaders ($MicrosoftNLB_NodeSetting |
                                                        Select-Object AliveMessagePeriod,AliveMessageTolerance,ClusterModeOnStart,ClusterModeSuspendOnStart,
                                                              DedicatedIPAddresses,DedicatedNetworkMasks,DescriptorsPerAlloc,
                                                              FilterIcmp,HostPriority,IpSecDescriptorTimeout,MaskSourceMAC,MaxConnectionDescriptors,MaxDescriptorsPerAlloc,Name,
                                                              NumActions,NumAliveMessages,NumberOfRules,NumPackets,PersistSuspendOnReboot,RemoteControlUDPPort,SettingID,TcpDescriptorTimeout)
                $MicrosoftNLB_PortRuleEx = RenameHeaders ($MicrosoftNLB_PortRuleEx | 
                                                        Select-Object StartPort,EndPort,Protocol,FilteringMode,Priority,LoadWeight,Affinity,ClientStickinessTimeout)
			}
            If ($MSCluster_Cluster -ne $null) {
				$MSCluster_Cluster = RenameHeaders ($MSCluster_Cluster | 
					Select-Object Name,QuorumType,QuorumPath,					
					@{n='Quorum Type Value';e={@("Unknown","Node","File Share Witness","Storage","None")[(($_.QuorumTypeValue)+1)]}},
					SharedVolumesRoot)
				$MSCluster_Node = $MSCluster_Node | 
	                Select-Object -Property Name,
						@{n='State';e={@("Unknown","Up","Down","Paused","Joining")[(($_.state)+1)]}} 
				$MSCluster_ResourceGroup = RenameHeaders ($MSCluster_ResourceGroup |
	                Select-Object -Property Name,
						@{n='State';e={@("Unknown","Online","Offline","Failed","Partial Online","Pending")[(($_.State)+1)]}},
						@{n='DefaultState';e={If($_.PersistentState -eq $true) {"Online"} else {"Offline"}}},
						@{n='PreferredOwner';e={WMIAssociationGrouptoPart $_.Name $MSCluster_ResourceGroupToPreferredNode }},						
						@{n='Resource';e={WMIAssociationGrouptoPart $_.Name $MSCluster_ResourceGroupToResource }},						
						@{n='Failback';e={@("No","Yes")[(($_.AutoFailbackType))]}})
				$MSCluster_Resource = RenameHeaders ($MSCluster_Resource | 
	                Select-Object -Property Name,
						@{n='State';e={@("Unknown","Inherited","Initializing","Online","Offline","Failed","Pending","Online Pending","Offline Pending")[(($_.State)+1)]}},
						@{n='RestartAction';e={@("Do not restart","Restart - do not failover","Restart - then failover")[(($_.RestartAction))]}},
						@{n='PossibleOwner';e={WMIAssociationGrouptoPart $_.Name $MSCluster_ResourceToPossibleOwner }},						
						RestartDelay,RestartPeriod,RestartThreshold)
				$MSCluster_Network = RenameHeaders ($MSCluster_Network | 
	                Select-Object -Property Name,
						Address,AddressMask,AutoMetric,
						@{n='Role';e={@("None","Cluster","Client","Client and Client")[($_.role)]}} | sort name)
				$MSCluster_DiskPartition = RenameHeaders ($MSCluster_DiskPartition | 
	                Select-Object -Property @{n='Resource';e={WMIAssociationPartToGroup (WMIAssociationPartToGroup $_.path $MSCluster_DiskToDiskPartition) $MSCluster_ResourceToDisk }},
						Path,PartitionNumber,VolumeLabel,FileSystem,Flags,
                    	@{n='FreeSpace';e={"{0:N0}" -f ($_.FreeSpace) + " MB"}},
                    	@{n='TotalSize';e={"{0:N0}" -f ($_.TotalSize) + " MB"}} | sort Resource)
			}
        }
    }                       
    $Date = Get-Date -format d
    $percent = 0
        switch ($Output.tolower())
        {
            "html" {
                $menu = ""
                $Header = @()
                $Footer = @()
                $Body = @()
                $Header += "<H1>Documentation For " + $Win32_ComputerSystem.Name + "</H1><BR>"       
                $Header += "<H3>Collected " + (get-date).ToString((Get-culture).DateTimeFormat.FullDateTimePattern) + "</H3><BR>"        
                $Footer += "<H3>Created by PowerSYDI, version " + $Version + "</H3><BR>"       
                $Body += "<BR>"

                $menu = "<li><a href=""#Computer System"">System</a></li>"
                $Body += OutputHTML 2 "Computer System" $false Win32_ComputerSystem 
                $Body += OutputHTML 2 "Operating System" $false Win32_OperatingSystem 
                       
                if ($ReportHardware) {
                    $menu += "<li><a href=""#Hardware"">Hardware</a></li>"
                    $Body += "<BR><H2><a name=""Hardware"">Hardware</a></H2><BR>" 
                    $Body += OutputHTML 3 "Processor" $true "Win32_Processor"
                    $Body += OutputHTML 3 "BIOS" $false "Win32_BIOS"
                    $Body += OutputHTML 3 "Computer System" $true "Win32_ComputerSystemProduct"
                    $Body += OutputHTML 3 "System Enclosure" $true "Win32_SystemEnclosure"
                    $Body += OutputHTML 3 "Physical Memory" $true "Win32_PhysicalMemory"
                    $Body += OutputHTML 3 "CDROM Drive" $true "Win32_CDROMDrive"
                    $Body += OutputHTML 3 "Sound Device" $true "Win32_SoundDevice"
                    $Body += OutputHTML 3 "Graphics" $true "Win32_VideoController"
                    $Body += OutputHTML 3 "Tape Drive Device" $true "Win32_TapeDrive"
                    $Body += OutputHTML 3 "Motherboard" $true "Win32_BaseBoard"
                    $Body += OutputHTML 3 "Keyboard" $true "Win32_Keyboard"
                    $Body += OutputHTML 3 "Pointing Device" $true "Win32_PointingDevice"
                }             
                if ($ReportSoftware) {
                    $menu += "<li><a href=""#Software"">Software</a></li>"
                    $Body += "<BR><H2><a name=""Software"">Software</a></H2><BR>" 
                    $Body += OutputHTML 3 "Software" $true "Win32_Product"
                    $Body += OutputHTML 3 "Patches" $true "Win32_QuickFixEngineering"
                    $Body += OutputHTML 3 "Optional Feature" $true "Win32_OptionalFeature"
                }             
                if ($ReportStorage) {
                    $menu += "<li><a href=""#Storage"">Storage</a></li>"
                    $Body += "<BR><H2><a name=""Storage"">Storage</a></H2><BR>" 
                    $Body += OutputHTML 3 "Disks" $false "Disks"
                }             
                if ($ReportNetwork) {
                    $menu += "<li><a href=""#Network"">Network</a></li>"
                    $Body += "<BR><H2><a name=""Network"">Network</a></H2><BR>" 
                    $Body += OutputHTML 3 "NICs" $false "Win32_NetworkAdapterConfiguration"
                    $Body += OutputHTML 3 "IP4 Route Table" $true "Win32_IP4RouteTable"
                }
                if ($ReportUser) {
                    $menu += "<li><a href=""#Users"">Users</a></li>"
                    $Body += "<BR><H2><a name=""Users"">User Information</a></H2><BR>" 
                    $Body += OutputHTML 3 "Users" $true "Win32_UserAccount"
                    $Body += OutputHTML 3 "Groups" $true "Win32_Group"
                }                 
                if ($ReportMisc) {
                    $menu += "<li><a href=""#General"">General</a></li>"
                    $Body += "<BR><H2><a name=""General"">General Information</a></H2><BR>" 
                    $Body += OutputHTML 3 "Event Logs" $true "Win32_NTEventLogFile"
                    $Body += OutputHTML 3 "Share" $true "Win32_Share"
                    $Body += OutputHTML 3 "Services" $true "Win32_Service"
                    $Body += OutputHTML 3 "Registry" $true "Win32_Registry"
                    $Body += OutputHTML 3 "Printers" $true "Win32_Printer"
                    $Body += OutputHTML 3 "TCPIP Printer Port" $true Win32_TCPIPPrinterPort 
                    $Body += OutputHTML 3 "Printer Drivers" $true Win32_PrinterDriver                       
                    $Body += OutputHTML 3 "TimeZone" $true "Win32_TimeZone"
                    $Body += OutputHTML 3 "Page File" $true "Win32_PageFile"
                    $Body += OutputHTML 3 "Environment Variables" $true "Win32_Environment"
                    $Body += OutputHTML 3 "Domain Information" $true "Win32_NTDomain"                                                  
                    $Body += OutputHTML 3 "General System" $false "Win32_General"
                }                 
                if ($ReportExtra) {
                    $menu += "<li><a href=""#ExtraInfo"">Extra</a></li>"
                    $Body += "<BR><H2><a name=""ExtraInfo"">Extra Information</a></H2><BR>" 
                    $Body += OutputHTML 3 "NLB Cluster" $true "MicrosoftNLB_Cluster"
                    $Body += OutputHTML 3 "NLB Cluster Settings" $true "MicrosoftNLB_ClusterSetting"
                    $Body += OutputHTML 3 "NLB Cluster Node" $true "MicrosoftNLB_Node"
                    $Body += OutputHTML 3 "NLB Cluster Node Settings" $false "MicrosoftNLB_NodeSetting"
                    $Body += OutputHTML 3 "NLB Cluster Port Settings" $true "MicrosoftNLB_PortRuleEx"
					
                    $Body += OutputHTML 3 "MS Cluster" $true "MSCluster_Cluster"
					$Body += OutputHTML 3 "MS Cluster Node" $true "MSCluster_Node"
					$Body += OutputHTML 3 "MS Cluster Resource Group" $true "MSCluster_ResourceGroup"
					$Body += OutputHTML 3 "MS Cluster Resource" $true "MSCluster_Resource"
					
					$Body += OutputHTML 3 "MS Cluster Network" $true MSCluster_Network
					$Body += OutputHTML 3 "MS Cluster Disks" $true MSCluster_DiskPartition
				}

                $HTML = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">
                    <html xmlns=""http://www.w3.org/1999/xhtml"" xml:lang=""en"" lang=""en"">
                    <head><title>PowerSYDI - $Computer</title><meta http-equiv=""content-type"" content=""text/html; charset=utf-8"" />$style</head>
                    <body>
                    <div id=""header"">$header</div>
                    <div id=""leftColumn""><h2>Links</h2><ul>$menu</ul><!--//end #tags//--></div>
                    <div id=""centerColumn"">$Body</div>
                    </body>
                    <div id=""footer"">$footer</div></html>"


                $filename = $path + $win32_computersystem.name.tolower() + ".htm"
                $HTML | Out-File $filename 
                if ($LoadDoc) {
                    Invoke-Expression $filename 
                }
            }
        "xml" {
            $xmlobjects = "ComputerSystem=`$Win32_ComputerSystem;OperatingSystem=`$Win32_OperatingSystem"
            if ($ReportHardware) {
                $Win32Hardware | foreach {
                    $name = $_ -replace "Win32_",""
                    $xmlobjects += "; $name=`$$_" 
                }
            }             
            if ($ReportSoftware) {
                $Win32Software | foreach {
                    $name = $_ -replace "Win32_",""
                    $xmlobjects += "; $name=`$$_" 
                }           
            }             
            if ($ReportStorage) {
                $xmlobjects += ";Disks=`$Disks"
            }             
            if ($ReportNetwork) {
                $Win32Network | foreach {
                    $name = $_ -replace "Win32_",""
                    $xmlobjects += "; $name=`$$_" 
                }           
            }
            if ($ReportMisc) {
                $Win32Misc | foreach {
                    $name = $_ -replace "Win32_",""
                    $xmlobjects += "; $name=`$$_" 
                }           
            }

            $xmlcmd = "`$xml = [PSCustomObject]@{$xmlobjects}"
            Invoke-Expression $xmlcmd

            $filename = $path + $win32_computersystem.name.tolower() + ".xml"
            $xml | ConvertTo-Xml -As string -Depth 4 | Out-File $filename 
            if ($LoadDoc) {
                Invoke-Expression $filename 
            }                        
        }
    }
} else {
    $Output = ""
    Write-Host "Server unavailable or not responding to pings"
}

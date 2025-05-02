##########################################
#Created by Nigel Tatschner on 22/07/2013#
##########################################
<#
	.SYNOPSIS
		 To gather Information about a Single or Multiple Systems.

	.DESCRIPTION
		This cmdlet gathers info about a computer(s) Hostname, IP Addresses and Speed, Mac Address, OS Version, Build Number, Service Pack Version, OS Architecture, Processor Architecture and Last Logged on User.

	.PARAMETER  ComputerName
		Enter a IP address, DNS name or Array of the machines.
    
    .PARAMETER Credential
        Use this Parameter to supply alternative creds that will work on a remote machine.

	.EXAMPLE
		PS C:\> Get-NTSystemInfo -ComputerName IAMACOMPUTER
        This gathers info about a specific machine "IAMCOMPUTER"
		
	.EXAMPLE
		PS C:\> Get-NTSystemInfo -ComputerName (Get-Content -Path C:\FileWithAComputerList.txt)

        This get the system info from a list of machines in the txt file "FileWithAComputerList.txt".
    
    .EXAMPLE
        PS c:\> Get-NTSystemInfo -ComputerName RemoteMachine -Credential (Get-Credential)

        This gathers info about a specific machine " RemoteMachine" Using alternate Credentials.

	.NOTES
		This funcion contains one parameter -ComputerName that can be pipped to.

#>

Function Get-NTSystemInfo {

[CmdletBinding()]

Param(

    [Parameter(Mandatory=$false,
     ValueFromPipeline=$True,
     HelpMessage="Enter Computer name or IP address to query")]
    [String[]] $ComputerName = 'localhost',
    
    [Parameter(ValueFromPipeline=$True)]
    [Object]$Credential

)
BEGIN{}

PROCESS
{

if ($Credential){
foreach ($Computer in $ComputerName) {
        $OSInfo = Get-WmiObject -class Win32_OperatingSystem -ComputerName $Computer -Credential $Credential
        $NetHardwareInfo = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $Computer -Credential $Credential
        $HardwareInfo = Get-WmiObject -Class Win32_Processor -ComputerName $Computer -Credential $Credential
        $NetHardwareInfo1 = Get-WmiObject -Class Win32_NetworkAdapter -ComputerName $Computer -Credential $Credential
        $UserSystemInfo = Get-WmiObject  -class win32_computerSystem -ComputerName $Computer -Credential $Credential

$Props =  @{'Hostname'=$OSInfo.CSName;
            'OS Version'=$OSInfo.name;
            'Build Number'=$OSInfo.BuildNumber;
            'Service Pack'=$OSInfo.ServicePackMajorVersion;
            'IP Addresses'=$NetHardwareInfo  | Sort-Object -Property IPAddress | Where-Object -Property IPAddress -GT $NULL | Select-Object -ExpandProperty IPAddress;
            'Mac Addresses'=$NetHardwareInfo | Sort-Object -Property MACAddress | Where-Object -Property MACAddress -GT $NULL | Select-Object -ExpandProperty MACAddress;
			'Network Speed'=$NetHardwareInfo1 | Where-Object -Property Speed -GT $NULL | Select-Object -ExpandProperty Speed ;
			'OS Architecture'=$OSInfo.OSArchitecture;
			'Processor Architecture'=$HardwareInfo.DataWidth;
			'Logged-in User' = $UserSystemInfo.Username;
            }
            $Object = New-Object -TypeName PSObject -Property $Props
Write-Output $Object | Select-Object -Property Hostname,'Logged-in User','IP Addresses','Network Speed','Mac Addresses','OS Version','Build Number','Service Pack','OS Architecture','Processor Architecture'


}} else {
foreach ($Computer in $ComputerName) {
        $OSInfo = Get-WmiObject -class Win32_OperatingSystem -ComputerName $Computer
        $NetHardwareInfo = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $Computer
        $HardwareInfo = Get-WmiObject -Class Win32_Processor -ComputerName $Computer
        $NetHardwareInfo1 = Get-WmiObject -Class Win32_NetworkAdapter -ComputerName $Computer
        $UserSystemInfo = Get-WmiObject  -class win32_computerSystem -ComputerName $Computer
        
$Props =  @{'Hostname'=$OSInfo.CSName;
            'OS Version'=$OSInfo.name;
            'Build Number'=$OSInfo.BuildNumber;
            'Service Pack'=$OSInfo.ServicePackMajorVersion;
            'IP Addresses'=$NetHardwareInfo  | Sort-Object -Property IPAddress | Where-Object -Property IPAddress -GT $NULL | Select-Object -ExpandProperty IPAddress;
            'Mac Addresses'=$NetHardwareInfo | Sort-Object -Property MACAddress | Where-Object -Property MACAddress -GT $NULL | Select-Object -ExpandProperty MACAddress;
			'Network Speed'=$NetHardwareInfo1 | Where-Object -Property Speed -GT $NULL | Select-Object -ExpandProperty Speed ;
			'OS Architecture'=$OSInfo.OSArchitecture;
			'Processor Architecture'=$HardwareInfo.DataWidth;
			'Logged-in User' = $UserSystemInfo.Username;
           }
           }
            
$Object = New-Object -TypeName PSObject -Property $Props
Write-Output $Object | Select-Object -Property Hostname,'Logged-in User','IP Addresses','Network Speed','Mac Addresses','OS Version','Build Number','Service Pack','OS Architecture','Processor Architecture'

}
}
}
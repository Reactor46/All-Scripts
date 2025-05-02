function Complete-WACNSP2SVPNConfiguration {
<#

.SYNOPSIS
Download, Extract VPN Folder File & Configure Point to Site VPN

.DESCRIPTION
This script is used to download VPN Client, extract VPN Folder File & Configure Point to Site VPN

.ROLE
Administrators

#>
param(
    [Parameter(Mandatory = $true)]
    [String]
    $AccessToken,
    [Parameter(Mandatory = $true)]
    [String]
    $ClientID,
    [Parameter(Mandatory = $true)]
    [String]
    $VNetSubnets, #"10.8.0.0/24;11.8.0.0/26"
    [Parameter(Mandatory = $true)]
    [String]
    $Subscription,
    [Parameter(Mandatory = $true)]
    [String]
    $ResourceGroup,
    [Parameter(Mandatory = $true)]
    [String]
    $GatewayName,
    [Parameter(Mandatory = $true)]
    [String]
    $VirtualNetwork,
    [Parameter(Mandatory = $true)]
    [String]
    $AddressSpace,
    [Parameter(Mandatory = $true)]
    [String]
    $Location
   
)
#Function to log event
function Log-MyEvent($Message){
    Try {
        $eventLogName = "ANA-LOG"
        $eventID = Get-Random -Minimum -1 -Maximum 65535
        #Create WAC specific Event Source if not exists
        $logFileExists = Get-EventLog -list | Where-Object {$_.logdisplayname -eq $eventLogName} 
        if (!$logFileExists) {
            New-EventLog -LogName $eventLogName -Source $eventLogName
        }
        #Prepare Event Log content and Write Event Log
        Write-EventLog -LogName $eventLogName -Source $eventLogName -EntryType Information -EventID $eventID -Message $Message

        $result = "Success"
    }
    Catch [Exception] {
        $result = $_.Exception.Message
    }
}

Function Build-Vpn( 
[Parameter(Mandatory = $true)]
[string]$XmlFilePathBuild,
[Parameter(Mandatory = $true)]
[string]$ProfileNameBuild,
[Parameter(Mandatory = $true)]
[string]$VNetGatewayNameBuild
)
{
    Log-MyEvent -Message "VPN Client Build started"
    
    #Enabling SC Config on demand
    $scConfigResult=CMD /C "sc config dmwappushservice start=demand"

    $a = Test-Path $xmlFilePathBuild
    echo $a

    $ProfileXML = Get-Content $xmlFilePathBuild

    echo $XML

    $ProfileNameBuildEscaped = $ProfileNameBuild -replace ' ', '%20'

    $Version = 201606090004

    $ProfileXML = $ProfileXML -replace '<', '&lt;'
    $ProfileXML = $ProfileXML -replace '>', '&gt;'
    $ProfileXML = $ProfileXML -replace '"', '&quot;'

    $nodeCSPURI = './Vendor/MSFT/VPNv2'
    $namespaceName = "root\cimv2\mdm\dmmap"
    $className = "MDM_VPNv2_01"

    $session = New-CimSession

    try
    {
        $newInstance = New-Object Microsoft.Management.Infrastructure.CimInstance $className, $namespaceName
        $property = [Microsoft.Management.Infrastructure.CimProperty]::Create("ParentID", "$nodeCSPURI", 'String', 'Key')
        $newInstance.CimInstanceProperties.Add($property)
        $property = [Microsoft.Management.Infrastructure.CimProperty]::Create("InstanceID", "$ProfileNameBuildEscaped", 'String', 'Key')
        $newInstance.CimInstanceProperties.Add($property)
        $property = [Microsoft.Management.Infrastructure.CimProperty]::Create("ProfileXML", "$ProfileXML", 'String', 'Property')
        $newInstance.CimInstanceProperties.Add($property)

        $session.CreateInstance($namespaceName, $newInstance)
        Log-MyEvent -Message "VPN Client Build completed."

        #Delete from RegEdit
        Remove-ItemProperty -path HKLM:\Software\WAC\VNetGatewayNotConfigured -name $VNetGatewayNameBuild -ErrorAction SilentlyContinue
        Log-MyEvent -Message "Removed from VNetGatewayNotConfigured RegEdit"

        #Delete File & Folders

        $folderToDelete = Split-Path -Path $xmlFilePathBuild

        Remove-Item -path $folderToDelete -Force -Recurse -ErrorAction SilentlyContinue
        Remove-Item -path $folderToDelete'.zip' -Force -Recurse -ErrorAction SilentlyContinue

        Log-MyEvent -Message "Trying to connect to VPN....."
        #Connect to this VPN
        $vpnConnected = rasdial $ProfileNameBuild
        Log-MyEvent -Message "VPN Connection established successfully."

        $Message = "Created $ProfileNameBuild profile."
        
        return "success"
    }
    catch [Exception]
    {
        Log-MyEvent -Message "Error Occured during establishing VPN"
        $Message = "Unable to create $ProfileNameBuild profile: $_"
        Log-MyEvent -Message $Message
        
        return $_.Exception.Message
    }
}

#Main operation started
Log-MyEvent -Message "Starting Gateway -'$GatewayName'"

$azureRmModule = Get-Module AzureRM -ListAvailable | Microsoft.PowerShell.Utility\Select-Object -Property Name -ErrorAction SilentlyContinue
if (!$azureRmModule.Name) {
    Log-MyEvent -Message "AzureRM module Not Available. Installing AzureRM Module"
    $packageProvIntsalled = Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
    $armIntalled = Install-Module AzureRm -Force
    Log-MyEvent -Message "Installed AzureRM Module successfully"
} 
else
{
    Log-MyEvent -Message "AzureRM Module Available"
}

import-module AzureRm

Log-MyEvent -Message "Imported AzureRM Module successfully"

#Login into Azure
#$logInRes = Login-AzureRmAccount -AccessToken $AccessToken -AccountId $ClientID
Log-MyEvent -Message "Logging in and selecting subscription..."
#Select Subscription
#$selectSubRes = Select-AzureRmSubscription -SubscriptionId $Subscription
$selectSubRes = Add-AzureRmAccount -AccessToken $AccessToken -AccountId $ClientID -Subscription $Subscription
if($selectSubRes)
{
    Log-MyEvent -Message "Selected Subscription successfully"

    #Select Gateway and generate URL to download VPN Client
    $profile = New-AzureRmVpnClientConfiguration -ResourceGroupName $ResourceGroup -Name $GatewayName -AuthenticationMethod "EapTls"
    if($profile)
    {
        Log-MyEvent -Message "URL generated to download VPN Client"

        #Create a Temp Folder if not exists
        $tempPath = "C:\WAC-TEMP"
        if (!(Test-Path $tempPath)) {
            $TempfolderCreated = New-Item -Path $tempPath -ItemType directory
        }

        #Delete previously downloaded zip file and extracted folder (if any)
        if (Test-Path "$tempPath\$GatewayName.zip") {
                Log-MyEvent -Message "Previous zip file found. deleting it.."
                Remove-Item -path "$tempPath\$GatewayName.zip" -Force -Recurse -ErrorAction SilentlyContinue
                Log-MyEvent -Message "Previous zip file deleted successfully."
        }
        if (Test-Path "$tempPath\$GatewayName") {
            Log-MyEvent -Message "Previous extracted folder found. deleting it.."
            Remove-Item -path "$tempPath\$GatewayName" -Force -Recurse -ErrorAction SilentlyContinue
            Log-MyEvent -Message "Previous extracted folder deleted successfully."
        }
    
        #Download VPN Client and save into a local temp folder
        $output = "$tempPath\" + $GatewayName + ".zip"
        $downLoadUrl = Invoke-WebRequest -Uri $profile.VPNProfileSASUrl -OutFile $output
        Log-MyEvent -Message "VPN Client downloaded successfully"

        #Extract zip
        $DestinationFolder = "$tempPath\" + $GatewayName
        Expand-Archive $output -DestinationPath $DestinationFolder
        Log-MyEvent -Message "VPN Client extracted successfully"

        #Read VPN Setting from Generic folder
        [xml]$XmlDocument = Get-Content -Path $DestinationFolder/Generic/VpnSettings.xml
        $vpnDNSRecord = $XmlDocument.VpnProfile.VpnServer
        Log-MyEvent -Message "Fetched VPN DNS record from VpnSettings.xml file in Generic foder"

        #Create a new VPN Profile name - check uniqueness
        $randomNumber = Get-Random -Minimum -1 -Maximum 65535
        $newVpnProfileName ='WACVPN-' + $randomNumber + '.xml'
        $isVpnAailable = 0
        while($isVpnAailable -eq 0)
        {
            if(!(Get-VpnConnection -Name $newVpnProfileName.split(".")[0] -ErrorAction SilentlyContinue))
            {
                $isVpnAailable=1
            }
            else
            { 
                $randomNumber = Get-Random -Minimum -1 -Maximum 65535
                $newVpnProfileName = 'WACVPN-' + $randomNumber + '.xml'
                $isVpnAailable=0
            }
        }
        try
        {
            Log-MyEvent -Message "Finalized VPN profile unique name"

            $xml_Path = $DestinationFolder + '\' + $newVpnProfileName
 
            #Set RasMan RegEdit value to 1
            $rasManPath = "HKLM:\System\CurrentControlSet\Services\RasMan\IKEv2"
            if((get-item -Path $rasManPath -ErrorAction SilentlyContinue))
            {
                Set-ItemProperty -Path $rasManPath -Name DisableCertReqPayload -Value 1
            }
            Log-MyEvent -Message "Updated RasMan to 1 in RegEdit"

            # Create the XML File Tags
            $xmlWriter = New-Object System.XMl.XmlTextWriter($xml_Path, $Null)
            $xmlWriter.Formatting = 'Indented'
            $xmlWriter.Indentation = 1
            $XmlWriter.IndentChar = "`t"
            $xmlWriter.WriteStartDocument()
            $xmlWriter.WriteStartElement('VPNProfile')
            $xmlWriter.WriteEndElement()
            $xmlWriter.WriteEndDocument()
            $xmlWriter.Flush()
            $xmlWriter.Close()
            Log-MyEvent -Message "XML File creation started"

            #Creating Root Node -NativeProfile
            $xmlDoc = [System.Xml.XmlDocument](Get-Content $xml_Path);
            $siteCollectionNode = $xmlDoc.CreateElement("NativeProfile")
            $nodeCreation = $xmlDoc.SelectSingleNode("//VPNProfile").AppendChild($siteCollectionNode)
            $xmlDoc.Save($xml_Path)

            #Creating Node -Servers
            $xmlDoc = [System.Xml.XmlDocument](Get-Content $xml_Path);
            $siteCollectionNode = $xmlDoc.CreateElement("Servers")
            $nodeCreation = $xmlDoc.SelectSingleNode("//VPNProfile/NativeProfile").AppendChild($siteCollectionNode)

            #Adding VPN DNS Record
            $RootFolderTextNode = $siteCollectionNode.AppendChild($xmlDoc.CreateTextNode($vpnDNSRecord));
            $xmlDoc.Save($xml_Path)

            #Creating Native Protocolol Type
            $xmlDoc = [System.Xml.XmlDocument](Get-Content $xml_Path);
            $siteCollectionNode = $xmlDoc.CreateElement("NativeProtocolType")
            $nodeCreation = $xmlDoc.SelectSingleNode("//VPNProfile/NativeProfile").AppendChild($siteCollectionNode)

            #Adding IKEv2
            $RootFolderTextNode = $siteCollectionNode.AppendChild($xmlDoc.CreateTextNode("IKEv2"));
            $xmlDoc.Save($xml_Path)

            #Creating Authentication
            $xmlDoc = [System.Xml.XmlDocument](Get-Content $xml_Path);
            $siteCollectionNode = $xmlDoc.CreateElement("Authentication")
            $nodeCreation = $xmlDoc.SelectSingleNode("//VPNProfile/NativeProfile").AppendChild($siteCollectionNode)
            $xmlDoc.Save($xml_Path)
            $xmlDoc = [System.Xml.XmlDocument](Get-Content $xml_Path);
            $siteCollectionNode = $xmlDoc.CreateElement("MachineMethod")
            $nodeCreation = $xmlDoc.SelectSingleNode("//VPNProfile/NativeProfile/Authentication").AppendChild($siteCollectionNode)
            $RootFolderTextNode = $siteCollectionNode.AppendChild($xmlDoc.CreateTextNode("Certificate"));
            $xmlDoc.Save($xml_Path)

            #Creating RoutingPolicyType
            $xmlDoc = [System.Xml.XmlDocument](Get-Content $xml_Path);
            $siteCollectionNode = $xmlDoc.CreateElement("RoutingPolicyType")
            $nodeCreation = $xmlDoc.SelectSingleNode("//VPNProfile/NativeProfile").AppendChild($siteCollectionNode)
            $RootFolderTextNode = $siteCollectionNode.AppendChild($xmlDoc.CreateTextNode("SplitTunnel"));
            $xmlDoc.Save($xml_Path)

            #Creating DisableClassBasedDefaultRoute
            $xmlDoc = [System.Xml.XmlDocument](Get-Content $xml_Path);
            $siteCollectionNode = $xmlDoc.CreateElement("DisableClassBasedDefaultRoute")
            $nodeCreation = $xmlDoc.SelectSingleNode("//VPNProfile/NativeProfile").AppendChild($siteCollectionNode)
            $RootFolderTextNode = $siteCollectionNode.AppendChild($xmlDoc.CreateTextNode("true"));
            $xmlDoc.Save($xml_Path)

            #Create Route
            $xmlDoc = [System.Xml.XmlDocument](Get-Content $xml_Path);
            $siteCollectionNode = $xmlDoc.CreateElement("Route")
            $nodeCreation = $xmlDoc.SelectSingleNode("//VPNProfile").AppendChild($siteCollectionNode)
            $xmlDoc.Save($xml_Path)

            #Get VNet Subnets and populate Address and prefix
            $allVnetSubnets = $VNetSubnets.split(";")
            foreach ($currentSubnet in $allVnetSubnets) {
                $address = $currentSubnet.split("/")[0]
                $prefixSize = $currentSubnet.split("/")[1]

                $xmlDoc = [System.Xml.XmlDocument](Get-Content $xml_Path);
                $siteCollectionNode = $xmlDoc.CreateElement("Address")
                $nodeCreation = $xmlDoc.SelectSingleNode("//VPNProfile/Route").AppendChild($siteCollectionNode)
                $RootFolderTextNode = $siteCollectionNode.AppendChild($xmlDoc.CreateTextNode($address));

                $siteCollectionNode = $xmlDoc.CreateElement("PrefixSize")
                $nodeCreation = $xmlDoc.SelectSingleNode("//VPNProfile/Route").AppendChild($siteCollectionNode)
                $RootFolderTextNode = $siteCollectionNode.AppendChild($xmlDoc.CreateTextNode($prefixSize));
    
                $xmlDoc.Save($xml_Path)
            }

            #Create TrafficFilter
            $xmlDoc = [System.Xml.XmlDocument](Get-Content $xml_Path);
            $siteCollectionNode = $xmlDoc.CreateElement("TrafficFilter")
            $nodeCreation = $xmlDoc.SelectSingleNode("//VPNProfile").AppendChild($siteCollectionNode)
            $xmlDoc.Save($xml_Path)

            #Get VNet Subnets and populate Address and prefix
            $allVnetSubnets = $VNetSubnets.split(";")
            foreach ($currentSubnet in $allVnetSubnets) {
                $xmlDoc = [System.Xml.XmlDocument](Get-Content $xml_Path);
                $siteCollectionNode = $xmlDoc.CreateElement("RemoteAddressRanges")
                $nodeCreation = $xmlDoc.SelectSingleNode("//VPNProfile/TrafficFilter").AppendChild($siteCollectionNode)
                $RootFolderTextNode = $siteCollectionNode.AppendChild($xmlDoc.CreateTextNode($currentSubnet));
    
                $xmlDoc.Save($xml_Path)
            }

            #Creating AlwaysOn
            $xmlDoc = [System.Xml.XmlDocument](Get-Content $xml_Path);
            $siteCollectionNode = $xmlDoc.CreateElement("AlwaysOn")
            $nodeCreation = $xmlDoc.SelectSingleNode("//VPNProfile").AppendChild($siteCollectionNode)
            $RootFolderTextNode = $siteCollectionNode.AppendChild($xmlDoc.CreateTextNode("true"));
            $xmlDoc.Save($xml_Path)

            #Creating DeviceTunnel
            $xmlDoc = [System.Xml.XmlDocument](Get-Content $xml_Path);
            $siteCollectionNode = $xmlDoc.CreateElement("DeviceTunnel")
            $nodeCreation = $xmlDoc.SelectSingleNode("//VPNProfile").AppendChild($siteCollectionNode)
            $RootFolderTextNode = $siteCollectionNode.AppendChild($xmlDoc.CreateTextNode("true"));
            $xmlDoc.Save($xml_Path)

            #Creating RegisterDNS
            $xmlDoc = [System.Xml.XmlDocument](Get-Content $xml_Path);
            $siteCollectionNode = $xmlDoc.CreateElement("RegisterDNS")
            $nodeCreation = $xmlDoc.SelectSingleNode("//VPNProfile").AppendChild($siteCollectionNode)
            $RootFolderTextNode = $siteCollectionNode.AppendChild($xmlDoc.CreateTextNode("true"));
            $xmlDoc.Save($xml_Path)

            #Removing XML Declaration 
            (Get-Content $xml_Path -raw).Replace('<?xml version="1.0"?>', '') | Set-Content $xml_Path;
            Log-MyEvent -Message "XML File creation completed"

            $returnType = ""
            $returnMsg = ""
            #Building VPN Client

            $buildStatus = Build-Vpn -XmlFilePathBuild $xml_Path -ProfileNameBuild $newVpnProfileName.split(".")[0] -VNetGatewayNameBuild $GatewayName
            if($buildStatus -eq "success")
            {
                #Create Registry Key and add Value to it
                $vpnConfiguredRegEditPath="HKLM:\Software\WAC\VPNConfigured"
                if(!(get-item -Path $vpnConfiguredRegEditPath -ErrorAction SilentlyContinue))
                {
                    $regKeyCreated = New-Item -Path HKLM:\Software -Name WAC\VPNConfigured -Force
                }
                $regKeyValue = $Subscription + ':' + $ResourceGroup + ':' + $GatewayName+ ':' + $VirtualNetwork+ ':' + $AddressSpace+':'+ $Location
    
                #Delete the previous gateway entry if already exists
                $readAllRegEdit = Get-Item -path $vpnConfiguredRegEditPath
                Foreach($thisRegEdit in $readAllRegEdit.Property)
                {
                    $thisRegValue = Get-ItemPropertyValue -path $vpnConfiguredRegEditPath -name $thisRegEdit
                    if($thisRegValue.ToLower() -eq $regKeyValue.ToLower())
                    {
                            Log-MyEvent -Message "Found previous connection with this Gateway. Deleting it"
                            Remove-ItemProperty -path $vpnConfiguredRegEditPath -name $thisRegEdit
                            Remove-VpnConnection -Name $thisRegEdit -Force -ErrorAction SilentlyContinue
                    }
                }
     
                Set-ItemProperty -Path $vpnConfiguredRegEditPath -Name $newVpnProfileName.split(".")[0] -Value $regKeyValue
                Log-MyEvent -Message "Logged into RegEdit successfully"

                $returnType = "success"
                $returnMsg = ""
            }
            else
            {
                #Delete from RegEdit
                Remove-ItemProperty -path HKLM:\Software\WAC\VNetGatewayNotConfigured -name $GatewayName -ErrorAction SilentlyContinue
                Log-MyEvent -Message "Removed from VNetGatewayNotConfigured RegEdit"
                $returnType = "fail"
                $returnMsg = "Building VPN on target machine failed"
            }
        }
        Catch [Exception] {
           Log-MyEvent -Message "Error occured during downloading and building VPN client"
           Log-MyEvent -Message $_.Exception.Message
           #Delete from RegEdit
           Remove-ItemProperty -path HKLM:\Software\WAC\VNetGatewayNotConfigured -name $GatewayName -ErrorAction SilentlyContinue
           Log-MyEvent -Message "Removed from VNetGatewayNotConfigured RegEdit"
           $returnType = "fail"
           $returnMsg = $_.Exception.Message
        }
        Log-MyEvent -Message "Ending Building process for -'$GatewayName'"
        $myResponse = New-Object -TypeName psobject

        $myResponse | Add-Member -MemberType NoteProperty -Name 'Status' -Value $returnType -ErrorAction SilentlyContinue
        $myResponse | Add-Member -MemberType NoteProperty -Name 'Message' -Value $returnMsg -ErrorAction SilentlyContinue

        $myResponse
    }
    else
    {
    
        Log-MyEvent -Message "Error Downloading VPN Client"
        #Delete from RegEdit
        Remove-ItemProperty -path HKLM:\Software\WAC\VNetGatewayNotConfigured -name $GatewayName -ErrorAction SilentlyContinue
        Log-MyEvent -Message "Removed from VNetGatewayNotConfigured RegEdit"
        Log-MyEvent -Message "Ending Building process with error for -'$GatewayName'"
    }
}
else
{
   Log-MyEvent -Message "Error in subscription selection."
   #Delete from RegEdit
   Remove-ItemProperty -path HKLM:\Software\WAC\VNetGatewayNotConfigured -name $GatewayName -ErrorAction SilentlyContinue
   Log-MyEvent -Message "Removed from VNetGatewayNotConfigured RegEdit"
   Log-MyEvent -Message "Ending Building process with error for -'$GatewayName'"
}
}
## [END] Complete-WACNSP2SVPNConfiguration ##
function Disable-WACNSAzureRmContextAutosave {
<#

.SYNOPSIS
Disable AzureRm Context Auto save

.DESCRIPTION
This script is used to disable AzureRm Context Auto save

.ROLE
Administrators

#>
$azureRmModule = Get-Module AzureRM -ListAvailable | Microsoft.PowerShell.Utility\Select-Object -Property Name -ErrorAction SilentlyContinue
if (!$azureRmModule.Name) {   
    $packageProvIntsalled = Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
    $armIntalled = Install-Module AzureRm -Force   
} 
Disable-AzureRmContextAutosave
}
## [END] Disable-WACNSAzureRmContextAutosave ##
function Get-WACNSClientAddressSpace {
<#

.SYNOPSIS
Get Client Address Space

.DESCRIPTION
This script is used to get client address space

.ROLE
Readers

#>
$clientAddressSpace = ""
Try
{
    #Fetch the IP Address of the Machine. There might be many IP Addresses, Here first index is getting fetched
    $ip = get-WmiObject Win32_NetworkAdapterConfiguration | Where {$_.Ipaddress.length -gt 1}
    $cidr = (Get-NetIPAddress -IPAddress $ip.ipaddress[0]).PrefixLength
    $clientaddr = "127.0.0.1/32"

    function INT64-toIP() { 
      param ([int64]$int) 
      return (([math]::truncate($int/16777216)).tostring()+"."+([math]::truncate(($int%16777216)/65536)).tostring()+"."+([math]::truncate(($int%65536)/256)).tostring()+"."+([math]::truncate($int%256)).tostring() )
    } 

    if ($cidr){
        $ipaddr = [Net.IPAddress]::Parse($ip.ipaddress[0])
        $maskaddr = [Net.IPAddress]::Parse((INT64-toIP -int ([convert]::ToInt64(("1"*$cidr+"0"*(32-$cidr)),2))))
        $networkaddr = new-object net.ipaddress ($maskaddr.address -band $ipaddr.address)
        $clientAddressSpace = "$networkaddr/$cidr"
    }
}
Catch
{
    $clientAddressSpace = ""
}
$clientAddressSpace
}
## [END] Get-WACNSClientAddressSpace ##
function Get-WACNSNetworks {
<#

.SYNOPSIS
Gets the network ip configuration.

.DESCRIPTION
Gets the network ip configuration. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>
Import-Module NetAdapter
Import-Module NetTCPIP
Import-Module DnsClient

Set-StrictMode -Version 5.0
$ErrorActionPreference = 'SilentlyContinue'

# Get all net information
$netAdapter = Get-NetAdapter

# conditions used to select the proper ip address for that object modeled after ibiza method.
# We only want manual (set by user manually), dhcp (set up automatically with dhcp), or link (set from link address)
# fe80 is the prefix for link local addresses, so that is the format want if the suffix origin is link
# SkipAsSource -eq zero only grabs ip addresses with skipassource set to false so we only get the preffered ip address
$ipAddress = Get-NetIPAddress | Where-Object {($_.SuffixOrigin -eq 'Manual') -or ($_.SuffixOrigin -eq 'Dhcp') -or (($_.SuffixOrigin -eq 'Link') -and (($_.IPAddress.StartsWith('fe80:')) -or ($_.IPAddress.StartsWith('2001:'))))}

$netIPInterface = Get-NetIPInterface
$netRoute = Get-NetRoute -PolicyStore ActiveStore
$dnsServer = Get-DnsClientServerAddress

# Load in relevant net information by name
Foreach ($currentNetAdapter in $netAdapter) {
    $result = New-Object PSObject

    # Net Adapter information
    $result | Add-Member -MemberType NoteProperty -Name 'InterfaceAlias' -Value $currentNetAdapter.InterfaceAlias
    $result | Add-Member -MemberType NoteProperty -Name 'InterfaceIndex' -Value $currentNetAdapter.InterfaceIndex
    $result | Add-Member -MemberType NoteProperty -Name 'InterfaceDescription' -Value $currentNetAdapter.InterfaceDescription
    $result | Add-Member -MemberType NoteProperty -Name 'Status' -Value $currentNetAdapter.Status
    $result | Add-Member -MemberType NoteProperty -Name 'MacAddress' -Value $currentNetAdapter.MacAddress
    $result | Add-Member -MemberType NoteProperty -Name 'LinkSpeed' -Value $currentNetAdapter.LinkSpeed

    # Net IP Address information
    # Primary addresses are used for outgoing calls so SkipAsSource is false (0)
    # Should only return one if properly configured, but it is possible to set multiple, so collect all
    $primaryIPv6Addresses = $ipAddress | Where-Object {($_.InterfaceAlias -eq $currentNetAdapter.Name) -and ($_.AddressFamily -eq 'IPv6') -and ($_.SkipAsSource -eq 0)}
    if ($primaryIPv6Addresses) {
        $ipArray = New-Object System.Collections.ArrayList
        $linkLocalArray = New-Object System.Collections.ArrayList
        Foreach ($address in $primaryIPv6Addresses) {
            if ($address -ne $null -and $address.IPAddress -ne $null -and $address.IPAddress.StartsWith('fe80')) {
                $linkLocalArray.Add(($address.IPAddress, $address.PrefixLength)) > $null
            }
            else {
                $ipArray.Add(($address.IPAddress, $address.PrefixLength)) > $null
            }
        }
        $result | Add-Member -MemberType NoteProperty -Name 'PrimaryIPv6Address' -Value $ipArray
        $result | Add-Member -MemberType NoteProperty -Name 'LinkLocalIPv6Address' -Value $linkLocalArray
    }

    $primaryIPv4Addresses = $ipAddress | Where-Object {($_.InterfaceAlias -eq $currentNetAdapter.Name) -and ($_.AddressFamily -eq 'IPv4') -and ($_.SkipAsSource -eq 0)}
    if ($primaryIPv4Addresses) {
        $ipArray = New-Object System.Collections.ArrayList
        Foreach ($address in $primaryIPv4Addresses) {
            $ipArray.Add(($address.IPAddress, $address.PrefixLength)) > $null
        }
        $result | Add-Member -MemberType NoteProperty -Name 'PrimaryIPv4Address' -Value $ipArray
    }

    # Secondary addresses are not used for outgoing calls so SkipAsSource is true (1)
    # There will usually not be secondary addresses, but collect them just in case
    $secondaryIPv6Adresses = $ipAddress | Where-Object {($_.InterfaceAlias -eq $currentNetAdapter.Name) -and ($_.AddressFamily -eq 'IPv6') -and ($_.SkipAsSource -eq 1)}
    if ($secondaryIPv6Adresses) {
        $ipArray = New-Object System.Collections.ArrayList
        Foreach ($address in $secondaryIPv6Adresses) {
            $ipArray.Add(($address.IPAddress, $address.PrefixLength)) > $null
        }
        $result | Add-Member -MemberType NoteProperty -Name 'SecondaryIPv6Address' -Value $ipArray
    }

    $secondaryIPv4Addresses = $ipAddress | Where-Object {($_.InterfaceAlias -eq $currentNetAdapter.Name) -and ($_.AddressFamily -eq 'IPv4') -and ($_.SkipAsSource -eq 1)}
    if ($secondaryIPv4Addresses) {
        $ipArray = New-Object System.Collections.ArrayList
        Foreach ($address in $secondaryIPv4Addresses) {
            $ipArray.Add(($address.IPAddress, $address.PrefixLength)) > $null
        }
        $result | Add-Member -MemberType NoteProperty -Name 'SecondaryIPv4Address' -Value $ipArray
    }

    # Net IP Interface information
    $currentDhcpIPv4 = $netIPInterface | Where-Object {($_.InterfaceAlias -eq $currentNetAdapter.Name) -and ($_.AddressFamily -eq 'IPv4')}
    if ($currentDhcpIPv4) {
        $result | Add-Member -MemberType NoteProperty -Name 'DhcpIPv4' -Value $currentDhcpIPv4.Dhcp
        $result | Add-Member -MemberType NoteProperty -Name 'IPv4Enabled' -Value $true
    }
    else {
        $result | Add-Member -MemberType NoteProperty -Name 'IPv4Enabled' -Value $false
    }

    $currentDhcpIPv6 = $netIPInterface | Where-Object {($_.InterfaceAlias -eq $currentNetAdapter.Name) -and ($_.AddressFamily -eq 'IPv6')}
    if ($currentDhcpIPv6) {
        $result | Add-Member -MemberType NoteProperty -Name 'DhcpIPv6' -Value $currentDhcpIPv6.Dhcp
        $result | Add-Member -MemberType NoteProperty -Name 'IPv6Enabled' -Value $true
    }
    else {
        $result | Add-Member -MemberType NoteProperty -Name 'IPv6Enabled' -Value $false
    }

    # Net Route information
    # destination prefix for selected ipv6 address is always ::/0
    $currentIPv6DefaultGateway = $netRoute | Where-Object {($_.InterfaceAlias -eq $currentNetAdapter.Name) -and ($_.DestinationPrefix -eq '::/0')}
    if ($currentIPv6DefaultGateway) {
        $ipArray = New-Object System.Collections.ArrayList
        Foreach ($address in $currentIPv6DefaultGateway) {
            if ($address.NextHop) {
                $ipArray.Add($address.NextHop) > $null
            }
        }
        $result | Add-Member -MemberType NoteProperty -Name 'IPv6DefaultGateway' -Value $ipArray
    }

    # destination prefix for selected ipv4 address is always 0.0.0.0/0
    $currentIPv4DefaultGateway = $netRoute | Where-Object {($_.InterfaceAlias -eq $currentNetAdapter.Name) -and ($_.DestinationPrefix -eq '0.0.0.0/0')}
    if ($currentIPv4DefaultGateway) {
        $ipArray = New-Object System.Collections.ArrayList
        Foreach ($address in $currentIPv4DefaultGateway) {
            if ($address.NextHop) {
                $ipArray.Add($address.NextHop) > $null
            }
        }
        $result | Add-Member -MemberType NoteProperty -Name 'IPv4DefaultGateway' -Value $ipArray
    }

    # DNS information
    # dns server util code for ipv4 is 2
    $currentIPv4DnsServer = $dnsServer | Where-Object {($_.InterfaceAlias -eq $currentNetAdapter.Name) -and ($_.AddressFamily -eq 2)}
    if ($currentIPv4DnsServer) {
        $ipArray = New-Object System.Collections.ArrayList
        Foreach ($address in $currentIPv4DnsServer) {
            if ($address.ServerAddresses) {
                $ipArray.Add($address.ServerAddresses) > $null
            }
        }
        $result | Add-Member -MemberType NoteProperty -Name 'IPv4DNSServer' -Value $ipArray
    }

    # dns server util code for ipv6 is 23
    $currentIPv6DnsServer = $dnsServer | Where-Object {($_.InterfaceAlias -eq $currentNetAdapter.Name) -and ($_.AddressFamily -eq 23)}
    if ($currentIPv6DnsServer) {
        $ipArray = New-Object System.Collections.ArrayList
        Foreach ($address in $currentIPv6DnsServer) {
            if ($address.ServerAddresses) {
                $ipArray.Add($address.ServerAddresses) > $null
            }
        }
        $result | Add-Member -MemberType NoteProperty -Name 'IPv6DNSServer' -Value $ipArray
    }

    $adapterGuid = $currentNetAdapter.InterfaceGuid
    if ($adapterGuid) {
      $regPath = "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\$($adapterGuid)"
      $ipv4Properties = Get-ItemProperty $regPath
      if ($ipv4Properties -and $ipv4Properties.NameServer) {
        $result | Add-Member -MemberType NoteProperty -Name 'IPv4DnsManuallyConfigured' -Value $true
      } else {
        $result | Add-Member -MemberType NoteProperty -Name 'IPv4DnsManuallyConfigured' -Value $false
      }

      $regPath = "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters\Interfaces\$($adapterGuid)"
      $ipv6Properties = Get-ItemProperty $regPath
      if ($ipv6Properties -and $ipv6Properties.NameServer) {
        $result | Add-Member -MemberType NoteProperty -Name 'IPv6DnsManuallyConfigured' -Value $true
      } else {
        $result | Add-Member -MemberType NoteProperty -Name 'IPv6DnsManuallyConfigured' -Value $false
      }
    }

    $result
}

}
## [END] Get-WACNSNetworks ##
function Get-WACNSRootCertValue {
<#

.SYNOPSIS
Storing Root and Client certificate, and then generate certificate value

.DESCRIPTION
This script is used to Storing Root and Client certificate provided by users, and then generate certificate value

.ROLE
Administrators

#>
param(
    [Parameter(Mandatory = $true)]
    [String]
    $RootCertPath,
    [Parameter(Mandatory = $true)]
    [String]
    $ClientCertPath,
    [Parameter(Mandatory = $true)]
    [String]
    $Password
)
$certName = ""
$content = ""

#Import Root certificate to Localmachine Root
if (Test-Path $RootCertPath) {
    $rootCertImported = Import-Certificate -FilePath $RootCertPath -certstorelocation 'Cert:\LocalMachine\Root'
    $certName = $rootCertImported.Subject.Split('=')[1]
    $content = @(
		[System.Convert]::ToBase64String($rootCertImported.RawData, 'InsertLineBreaks')
    )
    #Removing uploaded root cert file
    Remove-Item -path $RootCertPath -Force -Recurse -ErrorAction SilentlyContinue
}

#Import Client certificate to Localmachine My
if (Test-Path $ClientCertPath) {

    $securePassword = ConvertTo-SecureString $Password -asplaintext -force 
    $clientCertImported = Import-PfxCertificate -FilePath $ClientCertPath -CertStoreLocation Cert:\LocalMachine\My -Password $securePassword
    
    #Removing uploaded client cert file
    Remove-Item -path $ClientCertPath -Force -Recurse -ErrorAction SilentlyContinue
}

if($clientCertImported)
{
	$result = New-Object System.Object
	$result | Add-Member -MemberType NoteProperty -Name 'RootCertName' -Value $certName
	$result | Add-Member -MemberType NoteProperty -Name 'Content' -Value $content
	$result
}
}
## [END] Get-WACNSRootCertValue ##
function Get-WACNSVNetGatewayNameFromRegEdit {
<#

.SYNOPSIS
Reading Virtual Network Gateway information from Event Log

.DESCRIPTION
This Script is used to read Virtual Network Gateway information from Event Log

.ROLE
Administrators

#>

function Return-Object($rawData,$keyName)
{
	#Preparing Result object
    $subscriptionID = $rawData.Split(":")[0]
    $resourceGroup = $rawData.Split(":")[1]
    $vNetGateway = $rawData.Split(":")[2]

    $result = New-Object System.Object
    $result | Add-Member -MemberType NoteProperty -Name 'SubscriptionID' -Value $subscriptionID
    $result | Add-Member -MemberType NoteProperty -Name 'ResourceGroup' -Value $resourceGroup
    $result | Add-Member -MemberType NoteProperty -Name 'VNetGateway' -Value $vNetGateway
    $result | Add-Member -MemberType NoteProperty -Name 'KeyName' -Value $keyName
    $result
}

#Fetching from RegEdit (Only available/not configured)
$regEditPath = "HKLM:\Software\WAC\VNetGatewayNotConfigured"
$regItems = Get-Item -path $regEditPath -ErrorAction SilentlyContinue
Foreach($regitem in $regItems.Property)
{
  $regValue = Get-ItemPropertyValue -path $regEditPath -name $regitem
  Return-Object $regValue $regitem
}
}
## [END] Get-WACNSVNetGatewayNameFromRegEdit ##
function Get-WACNSVPNGatewayStatus {
<#

.SYNOPSIS
Check if the same gateway record is available

.DESCRIPTION
This Script is used to check if the same gateway record is available

.ROLE
Readers

#>
Param(
    [Parameter(Mandatory = $true)]
    [string] $Subscription,
    [Parameter(Mandatory = $true)]
    [string] $ResourceGroup,
    [Parameter(Mandatory = $true)]
    [string] $VNetGateway
)

$result = $false

$regKeyValue = $Subscription + ":" + $ResourceGroup + ":" + $VNetGateway

$vpnConfiguredRegEditPath = "HKLM:\Software\WAC\VNetGatewayNotConfigured"
if((Get-Item -Path $vpnConfiguredRegEditPath -ErrorAction SilentlyContinue))
{
    #check previous gateway entry if already exists
    $readAllRegEdit = Get-Item -Path $vpnConfiguredRegEditPath
    $isRecordAvailable = "0"
    Foreach($thisRegEdit in $readAllRegEdit.Property)
    {
        $thisRegValue = Get-ItemPropertyValue -Path $vpnConfiguredRegEditPath -Name $thisRegEdit
        if($thisRegValue.ToLower() -eq $regKeyValue.ToLower())
        {
            $isRecordAvailable = "1"
        }
    }

    if($isRecordAvailable -eq "1")
    {
         $result = $true
    }
    else
    {
         $result = $false
    }

}
else
{
    $result = $false
}

$result
}
## [END] Get-WACNSVPNGatewayStatus ##
function Get-WACNSVpnConnections {
<#

.SYNOPSIS
Get VPN Connections

.DESCRIPTION
This script is used to List all VPN Connection by reading from Registration Key and
matching with machine connected P2S VPn and Return Details

.ROLE
Readers

#>
Try
{
    $allVpnConnections = Get-VpnConnection -ErrorAction SilentlyContinue
    #Get VPN Profile Names from Registration Key
    $regEditPath = "HKLM:\SOFTWARE\WAC\VPNConfigured"
    $regItems = Get-Item -path $regEditPath -ErrorAction SilentlyContinue
    Foreach($regitem in $regItems.Property)
    {
        #Check if VPN Connection is available or not
        $thisVpn = $allVpnConnections | Where-Object {$_.name -eq $regitem} -ErrorAction SilentlyContinue
        if($thisVpn)
        {
            $regValue = Get-ItemPropertyValue -path $regEditPath -name $regitem -ErrorAction SilentlyContinue

            if($regValue)
            {
                #Generating response
                $connectionName = $regitem
                $description = "Point to Site VPN to Azure Virtual Network '"+ $regValue.split(":")[3]+"'"
                $connectionStatus = $thisVpn.ConnectionStatus
                $tunnelType = $thisVpn.TunnelType
                $vNetGatewayAddress = $thisVpn.ServerAddress
                $subscription = $regValue.split(":")[0]
                $resourceGroup = $regValue.split(":")[1]
                $vNetGateway = $regValue.split(":")[2]
                $virtualNetwork = $regValue.split(":")[3]
                $localNetworkAddressSpace = $regValue.split(":")[4]
                $location = $regValue.split(":")[5]

                #Preparing Object
                $myResponse = New-Object -TypeName psobject
                $myResponse | Add-Member -MemberType NoteProperty -Name 'Name' -Value $connectionName -ErrorAction SilentlyContinue
                $myResponse | Add-Member -MemberType NoteProperty -Name 'Description' -Value $description -ErrorAction SilentlyContinue
                $myResponse | Add-Member -MemberType NoteProperty -Name 'ConnectionStatus' -Value $connectionStatus -ErrorAction SilentlyContinue
                $myResponse | Add-Member -MemberType NoteProperty -Name 'TunnelType' -Value $tunnelType -ErrorAction SilentlyContinue
                $myResponse | Add-Member -MemberType NoteProperty -Name 'VnetGatewayAddress' -Value $vNetGatewayAddress -ErrorAction SilentlyContinue
                $myResponse | Add-Member -MemberType NoteProperty -Name 'Subscription' -Value $subscription -ErrorAction SilentlyContinue
                $myResponse | Add-Member -MemberType NoteProperty -Name 'ResourceGroup' -Value $resourceGroup -ErrorAction SilentlyContinue
                $myResponse | Add-Member -MemberType NoteProperty -Name 'VnetGateway' -Value $vNetGateway -ErrorAction SilentlyContinue
                $myResponse | Add-Member -MemberType NoteProperty -Name 'VirtualNetwork' -Value $virtualNetwork -ErrorAction SilentlyContinue
                $myResponse | Add-Member -MemberType NoteProperty -Name 'LocalNetworkAddressSpace' -Value $localNetworkAddressSpace -ErrorAction SilentlyContinue
                $myResponse | Add-Member -MemberType NoteProperty -Name 'Location' -Value $location -ErrorAction SilentlyContinue


                $myResponse
            }
        }
    }
}
Catch [Exception]{
    $myResponse = "Failed"
    $myResponse
}

}
## [END] Get-WACNSVpnConnections ##
function New-WACNSLogMyEvent {
<#

.SYNOPSIS
Logging My Event in Event Log

.DESCRIPTION
Logging My Event in Event Log

.ROLE
Administrators

#>
param(
    [Parameter(Mandatory = $true)]
    [String]
    $LogMessage
)
#Function to log event
function Log-MyEvent($Message){
    Try {
        $eventLogName = "ANA-LOG"
        $eventID = Get-Random -Minimum -1 -Maximum 65535
        #Create WAC specific Event Source if not exists
        $logFileExists = Get-EventLog -list | Where-Object {$_.logdisplayname -eq $eventLogName} 
        if (!$logFileExists) {
            New-EventLog -LogName $eventLogName -Source $eventLogName
        }
        #Prepare Event Log content and Write Event Log
        Write-EventLog -LogName $eventLogName -Source $eventLogName -EntryType Information -EventID $eventID -Message $Message

        $result = "Success"
    }
    Catch [Exception] {
        $result = $_.Exception.Message
    }
}

Log-MyEvent -Message "$LogMessage" 

}
## [END] New-WACNSLogMyEvent ##
function New-WACNSRegEditNotConfigured {
<#

.SYNOPSIS
Writing Virtual Network Gateway information into Event Log

.DESCRIPTION
This Script is used to store newly created Virtual Network Gateway information into Event Log

.ROLE
Administrators

#>
Param(
    [Parameter(Mandatory = $true)]
    [string] $Subscription,
    [Parameter(Mandatory = $true)]
    [string] $ResourceGroup,
    [Parameter(Mandatory = $true)]
    [string] $VNetGateway
)

$result = ""
Try {
    
    #Create Registry Key and add Value to it
    if(!(get-item -Path HKLM:\Software\WAC\VNetGatewayNotConfigured -ErrorAction SilentlyContinue))
    {
        $regKeyCreated = New-Item -Path HKLM:\Software -Name WAC\VNetGatewayNotConfigured -Force
    }
    $regKeyValue = $Subscription + ":" + $ResourceGroup + ":" + $VNetGateway
    Set-ItemProperty -Path HKLM:\Software\WAC\VNetGatewayNotConfigured -Name $VNetGateway -Value $regKeyValue
    
    $result = "Success"
}
Catch {
    $result = "Failed"
}
$result
}
## [END] New-WACNSRegEditNotConfigured ##
function New-WACNSSelfSignedRootCertificate {
<#

.SYNOPSIS
Create a Self-Signed Root certificate & Client Certificate

.DESCRIPTION
This script creates a Self-Signed Root certificate
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>
param(
    [Parameter(Mandatory = $true)]
    [String]
    $VNetGatewayName
   
)
$content=""
Try
{
    #Finalizing name of the certificate
    $uniqueRootCertName = $VNetGatewayName+'-P2SRoot-'+(Get-Date -UFormat "%m%d%Y%H%M")
    $uniqueClientCertName = $VNetGatewayName+'-P2SClient-'+(Get-Date -UFormat "%m%d%Y%H%M")

    #Creating Root Certificate
    $myCert = New-SelfSignedCertificate -Type Custom -KeySpec Signature -Subject "CN=$uniqueRootCertName" -KeyExportPolicy Exportable -HashAlgorithm sha256 -KeyLength 2048 -CertStoreLocation "Cert:\LocalMachine\My" -KeyUsageProperty Sign -KeyUsage CertSign

    #Creating client certificate
    $myClientCert = New-SelfSignedCertificate -Type Custom -DnsName $uniqueClientCertName -KeySpec Signature -Subject "CN=$uniqueClientCertName" -KeyExportPolicy Exportable -HashAlgorithm sha256 -KeyLength 2048 -CertStoreLocation "Cert:\LocalMachine\My" -Signer $myCert -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.2") 

    #Create a Temp Folder if not exists
    $tempPath = "C:\WAC-TEMP"
    if (!(Test-Path $tempPath)) {
       $tempfolderCreated = New-Item -Path $tempPath -ItemType directory
    }

    #Moving Root certificate from 'Cert:\LocalMachine\My' to 'Cert:\LocalMachine\Root'
    $exportLocation = $tempPath+"\$uniqueRootCertName.cer"
    $certExported = Export-Certificate -cert $myCert -filepath $exportLocation
    $certImported = Import-Certificate -FilePath $exportLocation -certstorelocation 'cert:\LocalMachine\Root'

    #Get Base64 Certificate Content
    $content = @(
		[System.Convert]::ToBase64String($myCert.RawData, 'InsertLineBreaks')
    )

    #Deleting temporary exported certificate file
    if (Test-Path $exportLocation) {
         Remove-Item -path $exportLocation -Force -Recurse -ErrorAction SilentlyContinue
    }
}
Catch [Exception]
{
    $content = $_.Exception.Message
}

$Result = New-Object System.Object
$Result | Add-Member -MemberType NoteProperty -Name 'RootCertName' -Value $uniqueRootCertName
$Result | Add-Member -MemberType NoteProperty -Name 'Content' -Value $content
$Result

}
## [END] New-WACNSSelfSignedRootCertificate ##
function New-WACNSTempFolder {
<#

.SYNOPSIS
Create a Temporary Folder in C drive of Target server

.DESCRIPTION
This script creates a Temporary Folder in C drive of Target server

.ROLE
Administrators

#>
$tempPath = "C:\WAC-TEMP"
if (!(Test-Path $tempPath)) {
    $tempFolderCreated = New-Item -Path $tempPath -ItemType directory
}
$tempPath
}
## [END] New-WACNSTempFolder ##
function Remove-WACNSNotConfiguredGateway {
<#

.SYNOPSIS
Remove Not found gateway or not valid provisioning status of Gateway entry from VNetGatewayNotConfigured RegEdit

.DESCRIPTION
This script is used to Remove Not found gateway or not valid provisioning status of Gateway entry from VNetGatewayNotConfigured RegEdit

.ROLE
Administrators

#>
param(
    [Parameter(Mandatory = $true)]
    [String]
    $VNetGatewayName,
    [Parameter(Mandatory = $true)]
    [String]
    $TenantId,
    [Parameter(Mandatory = $true)]
    [String]
    $AppId
)
#Function to log event
function Log-MyEvent($Message) {
    Try {
        $eventLogName = "ANA-LOG"
        $eventID = Get-Random -Minimum -1 -Maximum 65535
        #Create WAC specific Event Source if not exists
        $logFileExists = Get-EventLog -list | Where-Object {$_.logdisplayname -eq $eventLogName} 
        if (!$logFileExists) {
            New-EventLog -LogName $eventLogName -Source $eventLogName
        }
        #Prepare Event Log content and Write Event Log
        Write-EventLog -LogName $eventLogName -Source $eventLogName -EntryType Information -EventID $eventID -Message $Message

        $result = "Success"
    }
    Catch [Exception] {
        $result = $_.Exception.Message
    }
}

Log-MyEvent -Message "Gateway $VNetGatewayName doesn't exists or in failed state. so deleting this Gateway. Directory ID- $TenantId and App Id- $AppId" 
Remove-ItemProperty -path HKLM:\Software\WAC\VNetGatewayNotConfigured -name $VNetGatewayName
Log-MyEvent -Message "Gateway $VNetGatewayName has been deleted" 
}
## [END] Remove-WACNSNotConfiguredGateway ##
function Remove-WACNSVpnConnection {
<#

.SYNOPSIS
Remove VPN Connection

.DESCRIPTION
This script is used to remove VPN Connection

.ROLE
Administrators

#>
param(
    [Parameter(Mandatory = $true)]
    [String]
    $ConnectionName
   
)
#Removing VPN Connection
Remove-VpnConnection -Name $ConnectionName -Force

#Removing Item from RegEdit
Remove-ItemProperty -path HKLM:\Software\WAC\VPNConfigured -name $ConnectionName
}
## [END] Remove-WACNSVpnConnection ##
function Set-WACNSDhcpIP {
<#

.SYNOPSIS
Sets configuration of the specified network interface to use DHCP and updates DNS settings.

.DESCRIPTION
Sets configuration of the specified network interface to use DHCP and updates DNS settings. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

param (
    [Parameter(Mandatory = $true)] [string] $interfaceIndex,
    [Parameter(Mandatory = $true)] [string] $addressFamily,
    [string] $preferredDNS,
    [string] $alternateDNS)

Import-Module NetTCPIP

$ErrorActionPreference = 'Stop'

$ipInterface = Get-NetIPInterface -InterfaceIndex $interfaceIndex -AddressFamily $addressFamily
$netIPAddress = Get-NetIPAddress -InterfaceIndex $interfaceIndex -AddressFamily $addressFamily -ErrorAction SilentlyContinue
if ($addressFamily -eq "IPv4") {
    $prefix = '0.0.0.0/0'
}
else {
    $prefix = '::/0'
}

$netRoute = Get-NetRoute -InterfaceIndex $interfaceIndex -DestinationPrefix $prefix -ErrorAction SilentlyContinue

# avoid extra work if dhcp already set up
if ($ipInterface.Dhcp -eq 'Disabled') {
    if ($netIPAddress) {
        $netIPAddress | Remove-NetIPAddress -Confirm:$false
    }
    if ($netRoute) {
        $netRoute | Remove-NetRoute -Confirm:$false
    }

    $ipInterface | Set-NetIPInterface -DHCP Enabled
}

# reset or configure dns servers
$interfaceAlias = $ipInterface.InterfaceAlias
if ($preferredDNS) {
    netsh.exe interface $addressFamily set dnsservers name="$interfaceAlias" source=static validate=yes address="$preferredDNS"
    if (($LASTEXITCODE -eq 0) -and $alternateDNS) {
        netsh.exe interface $addressFamily add dnsservers name="$interfaceAlias" validate=yes address="$alternateDNS"
    }
}
else {
    netsh.exe interface $addressFamily delete dnsservers name="$interfaceAlias" address=all
}

# captures exit code of netsh.exe
$LASTEXITCODE

}
## [END] Set-WACNSDhcpIP ##
function Set-WACNSP2SVPNStatus {
<#

.SYNOPSIS
Connect / Disconnect P2S VPN

.DESCRIPTION
This script is used to Connect / Disconnect P2S VPN

.ROLE
Administrators

#>
param(
    [Parameter(Mandatory = $true)]
    [String]
    $VpnProfileName,
    [Parameter(Mandatory = $true)]
    [Int]
    $StatusFlag
    #Flag "1" is to Connect VPN. Flaf "0" to disconnect VPN.
)
if($StatusFlag -eq 1)
{
    #Connect VPN
	$result = rasdial $VpnProfileName
	$result = [String] $result
}
Elseif($StatusFlag -eq 0)
{
    #Disconnect VPN
    $result =  rasdial $VpnProfileName /disconnect
    $result= [String] $result
}
else
{
    $result = "No flag provided. Use 1 to connect and 0 to disconnect"
}
$statusProperty = "success"
$contentProperty = $result

if($result -match 'error' -or $result -match 'unacceptable' -or $result -match 'not')
{
	$statusProperty = "error"
}

#Preparing response Object
$response = New-Object System.Object
$response | Add-Member -MemberType NoteProperty -Name 'status' -Value $statusProperty
$response | Add-Member -MemberType NoteProperty -Name 'content' -Value $contentProperty
$response
}
## [END] Set-WACNSP2SVPNStatus ##
function Set-WACNSStaticIP {
<#

.SYNOPSIS
Sets configuration of the specified network interface to use a static IP address and updates DNS settings.

.DESCRIPTION
Sets configuration of the specified network interface to use a static IP address and updates DNS settings. The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>
param (
    [Parameter(Mandatory = $true)] [string] $interfaceIndex,
    [Parameter(Mandatory = $true)] [string] $ipAddress,
    [Parameter(Mandatory = $true)] [string] $prefixLength,
    [string] $defaultGateway,
    [string] $preferredDNS,
    [string] $alternateDNS,
    [Parameter(Mandatory = $true)] [string] $addressFamily)

Import-Module NetTCPIP

Set-StrictMode -Version 5.0
$ErrorActionPreference = 'Stop'

$netIPAddress = Get-NetIPAddress -InterfaceIndex $interfaceIndex -AddressFamily $addressFamily -ErrorAction SilentlyContinue

if ($addressFamily -eq "IPv4") {
    $prefix = '0.0.0.0/0'
}
else {
    $prefix = '::/0'
}

$netRoute = Get-NetRoute -InterfaceIndex $interfaceIndex -DestinationPrefix $prefix -ErrorAction SilentlyContinue

if ($netIPAddress) {
    $netIPAddress | Remove-NetIPAddress -Confirm:$false
}
if ($netRoute) {
    $netRoute | Remove-NetRoute -Confirm:$false
}

Set-NetIPInterface -InterfaceIndex $interfaceIndex -AddressFamily $addressFamily -DHCP Disabled

try {
    # this will fail if input is invalid
    if ($defaultGateway) {
        $netIPAddress | New-NetIPAddress -IPAddress $ipAddress -PrefixLength $prefixLength -DefaultGateway $defaultGateway -AddressFamily $addressFamily -ErrorAction Stop
    }
    else {
        $netIPAddress | New-NetIPAddress -IPAddress $ipAddress -PrefixLength $prefixLength -AddressFamily $addressFamily -ErrorAction Stop
    }
}
catch {
    # restore net route and ip address to previous values
    if ($netRoute -and $netIPAddress) {
        $netIPAddress | New-NetIPAddress -DefaultGateway $netRoute.NextHop -PrefixLength $netIPAddress.PrefixLength
    }
    elseif ($netIPAddress) {
        $netIPAddress | New-NetIPAddress
    }
    throw
}

$interfaceAlias = $netIPAddress.InterfaceAlias
if ($preferredDNS) {
    netsh.exe interface $addressFamily set dnsservers name="$interfaceAlias" source=static validate=yes address="$preferredDNS"
    if (($LASTEXITCODE -eq 0) -and $alternateDNS) {
        netsh.exe interface $addressFamily add dnsservers name="$interfaceAlias" validate=yes address="$alternateDNS"
    }
    return $LASTEXITCODE
}
else {
    return 0
}



}
## [END] Set-WACNSStaticIP ##
function Add-WACNSAdministrators {
<#

.SYNOPSIS
Adds administrators

.DESCRIPTION
Adds administrators

.ROLE
Administrators

#>

Param (
    [Parameter(Mandatory=$true)]
    [String] $usersListString
)


$usersToAdd = ConvertFrom-Json $usersListString
$adminGroup = Get-LocalGroup | Where-Object SID -eq 'S-1-5-32-544'

Add-LocalGroupMember -Group $adminGroup -Member $usersToAdd

Register-DnsClient -Confirm:$false

}
## [END] Add-WACNSAdministrators ##
function Disconnect-WACNSAzureHybridManagement {
<#

.SYNOPSIS
Disconnects a machine from azure hybrid agent.

.DESCRIPTION
Disconnects a machine from azure hybrid agent and uninstall the hybrid instance service.
The supported Operating Systems are Windows Server 2012 R2 and above.

.ROLE
Administrators

.PARAMETER tenantId
    The GUID that identifies a tenant in AAD

.PARAMETER authToken
    The authentication token for connection

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $tenantId,
    [Parameter(Mandatory = $true)]
    [String]
    $authToken
)

Set-StrictMode -Version 5.0

<#

.SYNOPSIS
Setup script runtime environment.

.DESCRIPTION
Setup script runtime environment.

#>

function setupScriptEnv() {
    Set-Variable -Name LogName -Option ReadOnly -Value "Microsoft-ServerManagementExperience" -Scope Script
    Set-Variable -Name LogSource -Option ReadOnly -Value "SMEScript" -Scope Script
    Set-Variable -Name ScriptName -Option ReadOnly -Value "Disconnect-HybridManagement.ps1" -Scope Script
    Set-Variable -Name HybridAgentExecutable -Option ReadOnly -Value "$env:ProgramFiles\AzureConnectedMachineAgent\azcmagent.exe" -Scope Script
    Set-Variable -Name HybridAgentPackage -Option ReadOnly -Value "Azure Connected Machine Agent" -Scope Script
}

<#

.SYNOPSIS
Cleanup script runtime environment.

.DESCRIPTION
Cleanup script runtime environment.

#>

function cleanupScriptEnv() {
    Remove-Variable -Name LogName -Scope Script -Force
    Remove-Variable -Name LogSource -Scope Script -Force
    Remove-Variable -Name ScriptName -Scope Script -Force
    Remove-Variable -Name HybridAgentExecutable -Scope Script -Force
    Remove-Variable -Name HybridAgentPackage -Scope Script -Force
}

<#

.SYNOPSIS
The main function.

.DESCRIPTION
Disconnects a machine from azure hybrid agent.

#>

function main(
    [string]$tenantId,
    [string]$authToken
) {
    $err = $null
    $args = @{}

   # Disconnect Azure hybrid agent
   & $HybridAgentExecutable disconnect --access-token $authToken

   # Uninstall Azure hybrid instance metadata service
   Uninstall-Package -Name $HybridAgentPackage -ErrorAction SilentlyContinue -ErrorVariable +err

   if ($err) {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
        -Message "[$ScriptName]:Could not uninstall the package. Error: $err"  -ErrorAction SilentlyContinue

        Write-Error @($err)[0]
        throw $err
   }

}


###############################################################################
# Script execution starts here
###############################################################################
setupScriptEnv

try {
    Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

    return main $tenantId $authToken

    return @()
} finally {
    cleanupScriptEnv
}

}
## [END] Disconnect-WACNSAzureHybridManagement ##
function Get-WACNSAzureHybridManagementConfiguration {
<#

.SYNOPSIS
Script that return the hybrid management configurations.

.DESCRIPTION
Script that return the hybrid management configurations.

.ROLE
Administrators

#>

Set-StrictMode -Version 5.0
Import-Module Microsoft.PowerShell.Management

<#

.SYNOPSIS
Setup script runtime environment.

.DESCRIPTION
Onboards a machine for hybrid management.

#>

function setupScriptEnv() {
    Set-Variable -Name LogName -Option ReadOnly -Value "Microsoft-ServerManagementExperience" -Scope Script
    Set-Variable -Name LogSource -Option ReadOnly -Value "SMEScript" -Scope Script
    Set-Variable -Name ScriptName -Option ReadOnly -Value "Get-HybridManagementConfiguration.ps1" -Scope Script
    Set-Variable -Name HybridAgentExecutable -Option ReadOnly -Value "$env:ProgramFiles\AzureConnectedMachineAgent\azcmagent.exe" -Scope Script
}

<#

.SYNOPSIS
Cleanup script runtime environment.

.DESCRIPTION
Cleanup script runtime environment.

#>

function cleanupScriptEnv() {
    Remove-Variable -Name LogName -Scope Script -Force
    Remove-Variable -Name LogSource -Scope Script -Force
    Remove-Variable -Name ScriptName -Scope Script -Force
    Remove-Variable -Name HybridAgentExecutable -Scope Script -Force
}

function main() {
    $config = & $HybridAgentExecutable show

    if ($config -and $config.count -gt 10) {
        @{ 
            machine = getValue($config[0]);
            resourceGroup = getValue($config[1]);
            subscriptionId = getValue($config[3]);
            tenantId = getValue($config[4])
            vmId = getValue($config[5]);
            azureRegion = getValue($config[7]);
            agentVersion = getValue($config[10]);
            agentStatus = getValue($config[12]);
            agentLastHeartbeat = getValue($config[13]);
        }
    } else {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
        -Message "[$ScriptName]:Could not find the Azure hybrid agent configuration."  -ErrorAction SilentlyContinue

        Write-Error @($err)[0]
        return @()
    }
}

function getValue([string]$keyValue) {
    $splitArray = $keyValue -split " : "
    $value = $splitArray[1].trim()
    return $value
}

###############################################################################
# Script execution starts here
###############################################################################
setupScriptEnv

try {
    Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

    return main

} finally {
    cleanupScriptEnv
}
}
## [END] Get-WACNSAzureHybridManagementConfiguration ##
function Get-WACNSAzureHybridManagementOnboardState {
<#

.SYNOPSIS
Script that returns if Azure Hybrid Agent is running or not.

.DESCRIPTION
Script that returns if Azure Hybrid Agent is running or not.

.ROLE
Readers

#>

Import-Module Microsoft.PowerShell.Management

$status = Get-Service -Name himds -ErrorAction SilentlyContinue
if ($null -eq $status) {
    # which means no such service is found.
    @{ Installed = $false; Running = $false }
}
elseif ($status.Status -eq "Running") {
    @{ Installed = $true; Running = $true }
}
else {
    @{ Installed = $true; Running = $false }
}

}
## [END] Get-WACNSAzureHybridManagementOnboardState ##
function Get-WACNSCimServiceDetail {
<#

.SYNOPSIS
Gets services in details using MSFT_ServerManagerTasks class.

.DESCRIPTION
Gets services in details using MSFT_ServerManagerTasks class.

.ROLE
Readers

#>

##SkipCheck=true##

Param(
)

import-module CimCmdlets

Invoke-CimMethod -Namespace root/microsoft/windows/servermanager -ClassName MSFT_ServerManagerTasks -MethodName GetServerServiceDetail

}
## [END] Get-WACNSCimServiceDetail ##
function Get-WACNSCimSingleService {
<#

.SYNOPSIS
Gets the service instance of CIM Win32_Service class.

.DESCRIPTION
Gets the service instance of CIM Win32_Service class.

.ROLE
Readers

#>

##SkipCheck=true##

Param(
[string]$Name
)

import-module CimCmdlets

$keyInstance = New-CimInstance -Namespace root/cimv2 -ClassName Win32_Service -Key @('Name') -Property @{Name=$Name;} -ClientOnly
Get-CimInstance $keyInstance

}
## [END] Get-WACNSCimSingleService ##
function Get-WACNSCimWin32LogicalDisk {
<#

.SYNOPSIS
Gets Win32_LogicalDisk object.

.DESCRIPTION
Gets Win32_LogicalDisk object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_LogicalDisk

}
## [END] Get-WACNSCimWin32LogicalDisk ##
function Get-WACNSCimWin32NetworkAdapter {
<#

.SYNOPSIS
Gets Win32_NetworkAdapter object.

.DESCRIPTION
Gets Win32_NetworkAdapter object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_NetworkAdapter

}
## [END] Get-WACNSCimWin32NetworkAdapter ##
function Get-WACNSCimWin32PhysicalMemory {
<#

.SYNOPSIS
Gets Win32_PhysicalMemory object.

.DESCRIPTION
Gets Win32_PhysicalMemory object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_PhysicalMemory

}
## [END] Get-WACNSCimWin32PhysicalMemory ##
function Get-WACNSCimWin32Processor {
<#

.SYNOPSIS
Gets Win32_Processor object.

.DESCRIPTION
Gets Win32_Processor object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_Processor

}
## [END] Get-WACNSCimWin32Processor ##
function Get-WACNSClusterInventory {
<#

.SYNOPSIS
Retrieves the inventory data for a cluster.

.DESCRIPTION
Retrieves the inventory data for a cluster.

.ROLE
Readers

#>

Import-Module CimCmdlets -ErrorAction SilentlyContinue

# JEA code requires to pre-import the module (this is slow on failover cluster environment.)
Import-Module FailoverClusters -ErrorAction SilentlyContinue

Import-Module Storage -ErrorAction SilentlyContinue
<#

.SYNOPSIS
Get the name of this computer.

.DESCRIPTION
Get the best available name for this computer.  The FQDN is preferred, but when not avaialble
the NetBIOS name will be used instead.

#>

function getComputerName() {
    $computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object Name, DNSHostName

    if ($computerSystem) {
        $computerName = $computerSystem.DNSHostName

        if ($null -eq $computerName) {
            $computerName = $computerSystem.Name
        }

        return $computerName
    }

    return $null
}

<#

.SYNOPSIS
Are the cluster PowerShell cmdlets installed on this server?

.DESCRIPTION
Are the cluster PowerShell cmdlets installed on this server?

#>

function getIsClusterCmdletAvailable() {
    $cmdlet = Get-Command "Get-Cluster" -ErrorAction SilentlyContinue

    return !!$cmdlet
}

<#

.SYNOPSIS
Get the MSCluster Cluster CIM instance from this server.

.DESCRIPTION
Get the MSCluster Cluster CIM instance from this server.

#>
function getClusterCimInstance() {
    $namespace = Get-CimInstance -Namespace root/MSCluster -ClassName __NAMESPACE -ErrorAction SilentlyContinue

    if ($namespace) {
        return Get-CimInstance -Namespace root/mscluster MSCluster_Cluster -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object fqdn, S2DEnabled
    }

    return $null
}


<#

.SYNOPSIS
Determines if the current cluster supports Failover Clusters Time Series Database.

.DESCRIPTION
Use the existance of the path value of cmdlet Get-StorageHealthSetting to determine if TSDB
is supported or not.

#>
function getClusterPerformanceHistoryPath() {
    $storageSubsystem = Get-StorageSubSystem clus* -ErrorAction SilentlyContinue
    $storageHealthSettings = Get-StorageHealthSetting -InputObject $storageSubsystem -Name "System.PerformanceHistory.Path" -ErrorAction SilentlyContinue

    return $null -ne $storageHealthSettings
}

<#

.SYNOPSIS
Get some basic information about the cluster from the cluster.

.DESCRIPTION
Get the needed cluster properties from the cluster.

#>
function getClusterInfo() {
    $returnValues = @{}

    $returnValues.Fqdn = $null
    $returnValues.isS2DEnabled = $false
    $returnValues.isTsdbEnabled = $false

    $cluster = getClusterCimInstance
    if ($cluster) {
        $returnValues.Fqdn = $cluster.fqdn
        $isS2dEnabled = !!(Get-Member -InputObject $cluster -Name "S2DEnabled") -and ($cluster.S2DEnabled -eq 1)
        $returnValues.isS2DEnabled = $isS2dEnabled

        if ($isS2DEnabled) {
            $returnValues.isTsdbEnabled = getClusterPerformanceHistoryPath
        } else {
            $returnValues.isTsdbEnabled = $false
        }
    }

    return $returnValues
}

<#

.SYNOPSIS
Are the cluster PowerShell Health cmdlets installed on this server?

.DESCRIPTION
Are the cluster PowerShell Health cmdlets installed on this server?

s#>
function getisClusterHealthCmdletAvailable() {
    $cmdlet = Get-Command -Name "Get-HealthFault" -ErrorAction SilentlyContinue

    return !!$cmdlet
}
<#

.SYNOPSIS
Are the Britannica (sddc management resources) available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) available on the cluster?

#>
function getIsBritannicaEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_Cluster -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Are the Britannica (sddc management resources) virtual machine available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) virtual machine available on the cluster?

#>
function getIsBritannicaVirtualMachineEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_VirtualMachine -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Are the Britannica (sddc management resources) virtual switch available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) virtual switch available on the cluster?

#>
function getIsBritannicaVirtualSwitchEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_VirtualSwitch -ErrorAction SilentlyContinue)
}

###########################################################################
# main()
###########################################################################

$clusterInfo = getClusterInfo

$result = New-Object PSObject

$result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $clusterInfo.Fqdn
$result | Add-Member -MemberType NoteProperty -Name 'IsS2DEnabled' -Value $clusterInfo.isS2DEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsTsdbEnabled' -Value $clusterInfo.isTsdbEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsClusterHealthCmdletAvailable' -Value (getIsClusterHealthCmdletAvailable)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaEnabled' -Value (getIsBritannicaEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaVirtualMachineEnabled' -Value (getIsBritannicaVirtualMachineEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaVirtualSwitchEnabled' -Value (getIsBritannicaVirtualSwitchEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsClusterCmdletAvailable' -Value (getIsClusterCmdletAvailable)
$result | Add-Member -MemberType NoteProperty -Name 'CurrentClusterNode' -Value (getComputerName)

$result

}
## [END] Get-WACNSClusterInventory ##
function Get-WACNSClusterNodes {
<#

.SYNOPSIS
Retrieves the inventory data for cluster nodes in a particular cluster.

.DESCRIPTION
Retrieves the inventory data for cluster nodes in a particular cluster.

.ROLE
Readers

#>

import-module CimCmdlets

# JEA code requires to pre-import the module (this is slow on failover cluster environment.)
import-module FailoverClusters -ErrorAction SilentlyContinue

###############################################################################
# Constants
###############################################################################

Set-Variable -Name LogName -Option Constant -Value "Microsoft-ServerManagementExperience" -ErrorAction SilentlyContinue
Set-Variable -Name LogSource -Option Constant -Value "SMEScripts" -ErrorAction SilentlyContinue
Set-Variable -Name ScriptName -Option Constant -Value $MyInvocation.ScriptName -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Are the cluster PowerShell cmdlets installed?

.DESCRIPTION
Use the Get-Command cmdlet to quickly test if the cluster PowerShell cmdlets
are installed on this server.

#>

function getClusterPowerShellSupport() {
    $cmdletInfo = Get-Command 'Get-ClusterNode' -ErrorAction SilentlyContinue

    return $cmdletInfo -and $cmdletInfo.Name -eq "Get-ClusterNode"
}

<#

.SYNOPSIS
Get the cluster nodes using the cluster CIM provider.

.DESCRIPTION
When the cluster PowerShell cmdlets are not available fallback to using
the cluster CIM provider to get the needed information.

#>

function getClusterNodeCimInstances() {
    # Change the WMI property NodeDrainStatus to DrainStatus to match the PS cmdlet output.
    return Get-CimInstance -Namespace root/mscluster MSCluster_Node -ErrorAction SilentlyContinue | `
        Microsoft.PowerShell.Utility\Select-Object @{Name="DrainStatus"; Expression={$_.NodeDrainStatus}}, DynamicWeight, Name, NodeWeight, FaultDomain, State
}

<#

.SYNOPSIS
Get the cluster nodes using the cluster PowerShell cmdlets.

.DESCRIPTION
When the cluster PowerShell cmdlets are available use this preferred function.

#>

function getClusterNodePsInstances() {
    return Get-ClusterNode -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object DrainStatus, DynamicWeight, Name, NodeWeight, FaultDomain, State
}

<#

.SYNOPSIS
Use DNS services to get the FQDN of the cluster NetBIOS name.

.DESCRIPTION
Use DNS services to get the FQDN of the cluster NetBIOS name.

.Notes
It is encouraged that the caller add their approprate -ErrorAction when
calling this function.

#>

function getClusterNodeFqdn([string]$clusterNodeName) {
    return ([System.Net.Dns]::GetHostEntry($clusterNodeName)).HostName
}

<#

.SYNOPSIS
Writes message to event log as warning.

.DESCRIPTION
Writes message to event log as warning.

#>

function writeToEventLog([string]$message) {
    Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Warning `
        -Message $message  -ErrorAction SilentlyContinue
}

<#

.SYNOPSIS
Get the cluster nodes.

.DESCRIPTION
When the cluster PowerShell cmdlets are available get the information about the cluster nodes
using PowerShell.  When the cmdlets are not available use the Cluster CIM provider.

#>

function getClusterNodes() {
    $isClusterCmdletAvailable = getClusterPowerShellSupport

    if ($isClusterCmdletAvailable) {
        $clusterNodes = getClusterNodePsInstances
    } else {
        $clusterNodes = getClusterNodeCimInstances
    }

    $clusterNodeMap = @{}

    foreach ($clusterNode in $clusterNodes) {
        $clusterNodeName = $clusterNode.Name.ToLower()
        try 
        {
            $clusterNodeFqdn = getClusterNodeFqdn $clusterNodeName -ErrorAction SilentlyContinue
        }
        catch 
        {
            $clusterNodeFqdn = $clusterNodeName
            writeToEventLog "[$ScriptName]: The fqdn for node '$clusterNodeName' could not be obtained. Defaulting to machine name '$clusterNodeName'"
        }

        $clusterNodeResult = New-Object PSObject

        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'FullyQualifiedDomainName' -Value $clusterNodeFqdn
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'Name' -Value $clusterNodeName
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'DynamicWeight' -Value $clusterNode.DynamicWeight
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'NodeWeight' -Value $clusterNode.NodeWeight
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'FaultDomain' -Value $clusterNode.FaultDomain
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'State' -Value $clusterNode.State
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'DrainStatus' -Value $clusterNode.DrainStatus

        $clusterNodeMap.Add($clusterNodeName, $clusterNodeResult)
    }

    return $clusterNodeMap
}

###########################################################################
# main()
###########################################################################

getClusterNodes

}
## [END] Get-WACNSClusterNodes ##
function Get-WACNSDecryptedDataFromNode {
<#

.SYNOPSIS
Gets data after decrypting it on a node.

.DESCRIPTION
Decrypts data on node using a cached RSAProvider used during encryption within 3 minutes of encryption and returns the decrypted data.
This script should be imported or copied directly to other scripts, do not send the returned data as an argument to other scripts.

.PARAMETER encryptedData
Encrypted data to be decrypted (String).

.ROLE
Readers

#>
param (
  [Parameter(Mandatory = $true)]
  [String]
  $encryptedData
)

Set-StrictMode -Version 5.0

New-Variable -Name rsaProviderInstanceName -Value "RSA" -Option Constant

function DecryptDataWithJWKOnNode {
  if(Get-Variable -Scope Global -Name $rsaProviderInstanceName -EA SilentlyContinue) {
    $rsaProvider = (Get-Variable -Scope Global -Name $rsaProviderInstanceName).Value
    $decryptedBytes = $rsaProvider.Decrypt([Convert]::FromBase64String($encryptedData), [System.Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)
    return [System.Text.Encoding]::UTF8.GetString($decryptedBytes)
  }
  # If you copy this script directly to another, you can get rid of the throw statement and add custom error handling logic such as "Write-Error"
  throw [System.InvalidOperationException] "Password decryption failed. RSACryptoServiceProvider Instance not found"
}

}
## [END] Get-WACNSDecryptedDataFromNode ##
function Get-WACNSEncryptionJWKOnNode {
<#

.SYNOPSIS
Gets encrytion JSON web key from node.

.DESCRIPTION
Gets encrytion JSON web key from node.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

New-Variable -Name rsaProviderInstanceName -Value "RSA" -Option Constant

function Get-RSAProvider
{
    if(Get-Variable -Scope Global -Name $rsaProviderInstanceName -EA SilentlyContinue)
    {
        return (Get-Variable -Scope Global -Name $rsaProviderInstanceName).Value
    }

    $Global:RSA = New-Object System.Security.Cryptography.RSACryptoServiceProvider -ArgumentList 4096
    return $RSA
}

function Get-JsonWebKey
{
    $rsaProvider = Get-RSAProvider
    $parameters = $rsaProvider.ExportParameters($false)
    return [PSCustomObject]@{
        kty = 'RSA'
        alg = 'RSA-OAEP'
        e = [Convert]::ToBase64String($parameters.Exponent)
        n = [Convert]::ToBase64String($parameters.Modulus).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    }
}

$jwk = Get-JsonWebKey
ConvertTo-Json $jwk -Compress

}
## [END] Get-WACNSEncryptionJWKOnNode ##
function Get-WACNSServerInventory {
<#

.SYNOPSIS
Retrieves the inventory data for a server.

.DESCRIPTION
Retrieves the inventory data for a server.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

Import-Module CimCmdlets

Import-Module Storage -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Converts an arbitrary version string into just 'Major.Minor'

.DESCRIPTION
To make OS version comparisons we only want to compare the major and
minor version.  Build number and/os CSD are not interesting.

#>

function convertOsVersion([string]$osVersion) {
  [Ref]$parsedVersion = $null
  if (![Version]::TryParse($osVersion, $parsedVersion)) {
    return $null
  }

  $version = [Version]$parsedVersion.Value
  return New-Object Version -ArgumentList $version.Major, $version.Minor
}

<#

.SYNOPSIS
Determines if CredSSP is enabled for the current server or client.

.DESCRIPTION
Check the registry value for the CredSSP enabled state.

#>

function isCredSSPEnabled() {
  Set-Variable credSSPServicePath -Option Constant -Value "WSMan:\localhost\Service\Auth\CredSSP"
  Set-Variable credSSPClientPath -Option Constant -Value "WSMan:\localhost\Client\Auth\CredSSP"

  $credSSPServerEnabled = $false;
  $credSSPClientEnabled = $false;

  $credSSPServerService = Get-Item $credSSPServicePath -ErrorAction SilentlyContinue
  if ($credSSPServerService) {
    $credSSPServerEnabled = [System.Convert]::ToBoolean($credSSPServerService.Value)
  }

  $credSSPClientService = Get-Item $credSSPClientPath -ErrorAction SilentlyContinue
  if ($credSSPClientService) {
    $credSSPClientEnabled = [System.Convert]::ToBoolean($credSSPClientService.Value)
  }

  return ($credSSPServerEnabled -or $credSSPClientEnabled)
}

<#

.SYNOPSIS
Determines if the Hyper-V role is installed for the current server or client.

.DESCRIPTION
The Hyper-V role is installed when the VMMS service is available.  This is much
faster then checking Get-WindowsFeature and works on Windows Client SKUs.

#>

function isHyperVRoleInstalled() {
  $vmmsService = Get-Service -Name "VMMS" -ErrorAction SilentlyContinue

  return $vmmsService -and $vmmsService.Name -eq "VMMS"
}

<#

.SYNOPSIS
Determines if the Hyper-V PowerShell support module is installed for the current server or client.

.DESCRIPTION
The Hyper-V PowerShell support module is installed when the modules cmdlets are available.  This is much
faster then checking Get-WindowsFeature and works on Windows Client SKUs.

#>
function isHyperVPowerShellSupportInstalled() {
  # quicker way to find the module existence. it doesn't load the module.
  return !!(Get-Module -ListAvailable Hyper-V -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Determines if Windows Management Framework (WMF) 5.0, or higher, is installed for the current server or client.

.DESCRIPTION
Windows Admin Center requires WMF 5 so check the registey for WMF version on Windows versions that are less than
Windows Server 2016.

#>
function isWMF5Installed([string] $operatingSystemVersion) {
  Set-Variable Server2016 -Option Constant -Value (New-Object Version '10.0')   # And Windows 10 client SKUs
  Set-Variable Server2012 -Option Constant -Value (New-Object Version '6.2')

  $version = convertOsVersion $operatingSystemVersion
  if (-not $version) {
    # Since the OS version string is not properly formatted we cannot know the true installed state.
    return $false
  }

  if ($version -ge $Server2016) {
    # It's okay to assume that 2016 and up comes with WMF 5 or higher installed
    return $true
  }
  else {
    if ($version -ge $Server2012) {
      # Windows 2012/2012R2 are supported as long as WMF 5 or higher is installed
      $registryKey = 'HKLM:\SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine'
      $registryKeyValue = Get-ItemProperty -Path $registryKey -Name PowerShellVersion -ErrorAction SilentlyContinue

      if ($registryKeyValue -and ($registryKeyValue.PowerShellVersion.Length -ne 0)) {
        $installedWmfVersion = [Version]$registryKeyValue.PowerShellVersion

        if ($installedWmfVersion -ge [Version]'5.0') {
          return $true
        }
      }
    }
  }

  return $false
}

<#

.SYNOPSIS
Determines if the current usser is a system administrator of the current server or client.

.DESCRIPTION
Determines if the current usser is a system administrator of the current server or client.

#>
function isUserAnAdministrator() {
  return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}

<#

.SYNOPSIS
Get some basic information about the Failover Cluster that is running on this server.

.DESCRIPTION
Create a basic inventory of the Failover Cluster that may be running in this server.

#>
function getClusterInformation() {
  $returnValues = @{ }

  $returnValues.IsS2dEnabled = $false
  $returnValues.IsCluster = $false
  $returnValues.ClusterFqdn = $null
  $returnValues.IsBritannicaEnabled = $false

  $namespace = Get-CimInstance -Namespace root/MSCluster -ClassName __NAMESPACE -ErrorAction SilentlyContinue
  if ($namespace) {
    $cluster = Get-CimInstance -Namespace root/MSCluster -ClassName MSCluster_Cluster -ErrorAction SilentlyContinue
    if ($cluster) {
      $returnValues.IsCluster = $true
      $returnValues.ClusterFqdn = $cluster.Fqdn
      $returnValues.IsS2dEnabled = !!(Get-Member -InputObject $cluster -Name "S2DEnabled") -and ($cluster.S2DEnabled -gt 0)
      $returnValues.IsBritannicaEnabled = $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_Cluster -ErrorAction SilentlyContinue)
    }
  }

  return $returnValues
}

<#

.SYNOPSIS
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the passed in computer name.

.DESCRIPTION
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the passed in computer name.

#>
function getComputerFqdnAndAddress($computerName) {
  $hostEntry = [System.Net.Dns]::GetHostEntry($computerName)
  $addressList = @()
  foreach ($item in $hostEntry.AddressList) {
    $address = New-Object PSObject
    $address | Add-Member -MemberType NoteProperty -Name 'IpAddress' -Value $item.ToString()
    $address | Add-Member -MemberType NoteProperty -Name 'AddressFamily' -Value $item.AddressFamily.ToString()
    $addressList += $address
  }

  $result = New-Object PSObject
  $result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $hostEntry.HostName
  $result | Add-Member -MemberType NoteProperty -Name 'AddressList' -Value $addressList
  return $result
}

<#

.SYNOPSIS
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the current server or client.

.DESCRIPTION
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the current server or client.

#>
function getHostFqdnAndAddress($computerSystem) {
  $computerName = $computerSystem.DNSHostName
  if (!$computerName) {
    $computerName = $computerSystem.Name
  }

  return getComputerFqdnAndAddress $computerName
}

<#

.SYNOPSIS
Are the needed management CIM interfaces available on the current server or client.

.DESCRIPTION
Check for the presence of the required server management CIM interfaces.

#>
function getManagementToolsSupportInformation() {
  $returnValues = @{ }

  $returnValues.ManagementToolsAvailable = $false
  $returnValues.ServerManagerAvailable = $false

  $namespaces = Get-CimInstance -Namespace root/microsoft/windows -ClassName __NAMESPACE -ErrorAction SilentlyContinue

  if ($namespaces) {
    $returnValues.ManagementToolsAvailable = !!($namespaces | Where-Object { $_.Name -ieq "ManagementTools" })
    $returnValues.ServerManagerAvailable = !!($namespaces | Where-Object { $_.Name -ieq "ServerManager" })
  }

  return $returnValues
}

<#

.SYNOPSIS
Check the remote app enabled or not.

.DESCRIPTION
Check the remote app enabled or not.

#>
function isRemoteAppEnabled() {
  Set-Variable key -Option Constant -Value "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Terminal Server\\TSAppAllowList"

  $registryKeyValue = Get-ItemProperty -Path $key -Name fDisabledAllowList -ErrorAction SilentlyContinue

  if (-not $registryKeyValue) {
    return $false
  }
  return $registryKeyValue.fDisabledAllowList -eq 1
}

<#

.SYNOPSIS
Check the remote app enabled or not.

.DESCRIPTION
Check the remote app enabled or not.

#>

<#
c
.SYNOPSIS
Get the Win32_OperatingSystem information as well as current version information from the registry

.DESCRIPTION
Get the Win32_OperatingSystem instance and filter the results to just the required properties.
This filtering will make the response payload much smaller. Included in the results are current version
information from the registry

#>
function getOperatingSystemInfo() {
  $operatingSystemInfo = Get-CimInstance Win32_OperatingSystem | Microsoft.PowerShell.Utility\Select-Object csName, Caption, OperatingSystemSKU, Version, ProductType, OSType, LastBootUpTime, SerialNumber
  $currentVersion = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" | Microsoft.PowerShell.Utility\Select-Object CurrentBuild, UBR, DisplayVersion

  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name CurrentBuild -Value $currentVersion.CurrentBuild
  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name UpdateBuildRevision -Value $currentVersion.UBR
  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name DisplayVersion -Value $currentVersion.DisplayVersion

  return $operatingSystemInfo
}

<#

.SYNOPSIS
Get the Win32_ComputerSystem information

.DESCRIPTION
Get the Win32_ComputerSystem instance and filter the results to just the required properties.
This filtering will make the response payload much smaller.

#>
function getComputerSystemInfo() {
  return Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue | `
    Microsoft.PowerShell.Utility\Select-Object TotalPhysicalMemory, DomainRole, Manufacturer, Model, NumberOfLogicalProcessors, Domain, Workgroup, DNSHostName, Name, PartOfDomain, SystemFamily, SystemSKUNumber
}

<#

.SYNOPSIS
Get SMBIOS locally from the passed in machineName


.DESCRIPTION
Get SMBIOS locally from the passed in machine name

#>
function getSmbiosData($computerSystem) {
  <#
    Array of chassis types.
    The following list of ChassisTypes is copied from the latest DMTF SMBIOS specification.
    REF: https://www.dmtf.org/sites/default/files/standards/documents/DSP0134_3.1.1.pdf
  #>
  $ChassisTypes =
  @{
    1  = 'Other'
    2  = 'Unknown'
    3  = 'Desktop'
    4  = 'Low Profile Desktop'
    5  = 'Pizza Box'
    6  = 'Mini Tower'
    7  = 'Tower'
    8  = 'Portable'
    9  = 'Laptop'
    10 = 'Notebook'
    11 = 'Hand Held'
    12 = 'Docking Station'
    13 = 'All in One'
    14 = 'Sub Notebook'
    15 = 'Space-Saving'
    16 = 'Lunch Box'
    17 = 'Main System Chassis'
    18 = 'Expansion Chassis'
    19 = 'SubChassis'
    20 = 'Bus Expansion Chassis'
    21 = 'Peripheral Chassis'
    22 = 'Storage Chassis'
    23 = 'Rack Mount Chassis'
    24 = 'Sealed-Case PC'
    25 = 'Multi-system chassis'
    26 = 'Compact PCI'
    27 = 'Advanced TCA'
    28 = 'Blade'
    29 = 'Blade Enclosure'
    30 = 'Tablet'
    31 = 'Convertible'
    32 = 'Detachable'
    33 = 'IoT Gateway'
    34 = 'Embedded PC'
    35 = 'Mini PC'
    36 = 'Stick PC'
  }

  $list = New-Object System.Collections.ArrayList
  $win32_Bios = Get-CimInstance -class Win32_Bios
  $obj = New-Object -Type PSObject | Microsoft.PowerShell.Utility\Select-Object SerialNumber, Manufacturer, UUID, BaseBoardProduct, ChassisTypes, Chassis, SystemFamily, SystemSKUNumber, SMBIOSAssetTag
  $obj.SerialNumber = $win32_Bios.SerialNumber
  $obj.Manufacturer = $win32_Bios.Manufacturer
  $computerSystemProduct = Get-CimInstance Win32_ComputerSystemProduct
  if ($null -ne $computerSystemProduct) {
    $obj.UUID = $computerSystemProduct.UUID
  }
  $baseboard = Get-CimInstance Win32_BaseBoard
  if ($null -ne $baseboard) {
    $obj.BaseBoardProduct = $baseboard.Product
  }
  $systemEnclosure = Get-CimInstance Win32_SystemEnclosure
  if ($null -ne $systemEnclosure) {
    $obj.SMBIOSAssetTag = $systemEnclosure.SMBIOSAssetTag
  }
  $obj.ChassisTypes = Get-CimInstance Win32_SystemEnclosure | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty ChassisTypes
  $obj.Chassis = New-Object -TypeName 'System.Collections.ArrayList'
  $obj.ChassisTypes | ForEach-Object -Process {
    $obj.Chassis.Add($ChassisTypes[[int]$_])
  }
  $obj.SystemFamily = $computerSystem.SystemFamily
  $obj.SystemSKUNumber = $computerSystem.SystemSKUNumber
  $list.Add($obj) | Out-Null

  return $list

}

<#

.SYNOPSIS
Get the azure arc status information

.DESCRIPTION
Get the azure arc status information

#>
function getAzureArcStatus() {

  $LogName = "Microsoft-ServerManagementExperience"
  $LogSource = "SMEScript"
  $ScriptName = "Get-ServerInventory.ps1 - getAzureArcStatus()"

  Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

  Get-Service -Name himds -ErrorVariable Err -ErrorAction SilentlyContinue | Out-Null

  if (!!$Err) {

    $Err = "The Azure arc agent is not installed. Details: $Err"

    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Information `
    -Message "[$ScriptName]: $Err"  -ErrorAction SilentlyContinue

    $status = "NotInstalled"
  }
  else {
    $status = (azcmagent show --json | ConvertFrom-Json -ErrorAction Stop).status
  }

  return $status
}

<#

.SYNOPSIS
Gets an EnforcementMode that describes the system lockdown policy on this computer.

.DESCRIPTION
By checking the system lockdown policy, we can infer if PowerShell is in ConstrainedLanguage mode as a result of an enforced WDAC policy.
Note: $ExecutionContext.SessionState.LanguageMode should not be used within a trusted (by the WDAC policy) script context for this purpose because
the language mode returned would potentially not reflect the system-wide lockdown policy/language mode outside of the execution context.

#>
function getSystemLockdownPolicy() {
  return [System.Management.Automation.Security.SystemPolicy]::GetSystemLockdownPolicy().ToString()
}

<#

.SYNOPSIS
Determines if the operating system is HCI.

.DESCRIPTION
Using the operating system 'Caption' (which corresponds to the 'ProductName' registry key at HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion) to determine if a server OS is HCI.

#>
function isServerOsHCI([string] $operatingSystemCaption) {
  return $operatingSystemCaption -eq "Microsoft Azure Stack HCI"
}

###########################################################################
# main()
###########################################################################

$operatingSystem = getOperatingSystemInfo
$computerSystem = getComputerSystemInfo
$isAdministrator = isUserAnAdministrator
$fqdnAndAddress = getHostFqdnAndAddress $computerSystem
$hostname = [Environment]::MachineName
$netbios = $env:ComputerName
$managementToolsInformation = getManagementToolsSupportInformation
$isWmfInstalled = isWMF5Installed $operatingSystem.Version
$clusterInformation = getClusterInformation -ErrorAction SilentlyContinue
$isHyperVPowershellInstalled = isHyperVPowerShellSupportInstalled
$isHyperVRoleInstalled = isHyperVRoleInstalled
$isCredSSPEnabled = isCredSSPEnabled
$isRemoteAppEnabled = isRemoteAppEnabled
$smbiosData = getSmbiosData $computerSystem
$azureArcStatus = getAzureArcStatus
$systemLockdownPolicy = getSystemLockdownPolicy
$isHciServer = isServerOsHCI $operatingSystem.Caption

$result = New-Object PSObject
$result | Add-Member -MemberType NoteProperty -Name 'IsAdministrator' -Value $isAdministrator
$result | Add-Member -MemberType NoteProperty -Name 'OperatingSystem' -Value $operatingSystem
$result | Add-Member -MemberType NoteProperty -Name 'ComputerSystem' -Value $computerSystem
$result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $fqdnAndAddress.Fqdn
$result | Add-Member -MemberType NoteProperty -Name 'AddressList' -Value $fqdnAndAddress.AddressList
$result | Add-Member -MemberType NoteProperty -Name 'Hostname' -Value $hostname
$result | Add-Member -MemberType NoteProperty -Name 'NetBios' -Value $netbios
$result | Add-Member -MemberType NoteProperty -Name 'IsManagementToolsAvailable' -Value $managementToolsInformation.ManagementToolsAvailable
$result | Add-Member -MemberType NoteProperty -Name 'IsServerManagerAvailable' -Value $managementToolsInformation.ServerManagerAvailable
$result | Add-Member -MemberType NoteProperty -Name 'IsWmfInstalled' -Value $isWmfInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsCluster' -Value $clusterInformation.IsCluster
$result | Add-Member -MemberType NoteProperty -Name 'ClusterFqdn' -Value $clusterInformation.ClusterFqdn
$result | Add-Member -MemberType NoteProperty -Name 'IsS2dEnabled' -Value $clusterInformation.IsS2dEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaEnabled' -Value $clusterInformation.IsBritannicaEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsHyperVRoleInstalled' -Value $isHyperVRoleInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsHyperVPowershellInstalled' -Value $isHyperVPowershellInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsCredSSPEnabled' -Value $isCredSSPEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsRemoteAppEnabled' -Value $isRemoteAppEnabled
$result | Add-Member -MemberType NoteProperty -Name 'SmbiosData' -Value $smbiosData
$result | Add-Member -MemberType NoteProperty -Name 'AzureArcStatus' -Value $azureArcStatus
$result | Add-Member -MemberType NoteProperty -Name 'SystemLockdownPolicy' -Value $systemLockdownPolicy
$result | Add-Member -MemberType NoteProperty -Name 'IsHciServer' -Value $isHciServer

$result

}
## [END] Get-WACNSServerInventory ##
function Resolve-WACNSDNSName {
<#

.SYNOPSIS
Resolve VM Provisioning

.DESCRIPTION
Resolve VM Provisioning

.ROLE
Administrators

#>

Param
(
    [string] $computerName
)

$succeeded = $null
$count = 0;
$maxRetryTimes = 15 * 100 # 15 minutes worth of 10 second sleep times
while ($count -lt $maxRetryTimes)
{
  $resolved =  Resolve-DnsName -Name $computerName -ErrorAction SilentlyContinue

    if ($resolved)
    {
      $succeeded = $true
      break
    }

    $count += 1

    if ($count -eq $maxRetryTimes)
    {
        $succeeded = $false
    }

    Start-Sleep -Seconds 10
}

Write-Output @{ "succeeded" = $succeeded }

}
## [END] Resolve-WACNSDNSName ##
function Resume-WACNSCimService {
<#

.SYNOPSIS
Resume a service using CIM Win32_Service class.

.DESCRIPTION
Resume a service using CIM Win32_Service class.

.ROLE
Readers

#>

##SkipCheck=true##

Param(
[string]$Name
)

import-module CimCmdlets

$keyInstance = New-CimInstance -Namespace root/cimv2 -ClassName Win32_Service -Key @('Name') -Property @{Name=$Name;} -ClientOnly
Invoke-CimMethod $keyInstance -MethodName ResumeService

}
## [END] Resume-WACNSCimService ##
function Set-WACNSAzureHybridManagement {
<#

.SYNOPSIS
Onboards a machine for hybrid management.

.DESCRIPTION
Sets up a non-Azure machine to be used as a resource in Azure
The supported Operating Systems are Windows Server 2012 R2 and above.

.ROLE
Administrators

.PARAMETER subscriptionId
    The GUID that identifies subscription to Azure services

.PARAMETER resourceGroup
    The container that holds related resources for an Azure solution

.PARAMETER tenantId
    The GUID that identifies a tenant in AAD

.PARAMETER azureRegion
    The region in Azure where the service is to be deployed

.PARAMETER useProxyServer
    The flag to determine whether to use proxy server or not

.PARAMETER proxyServerIpAddress
    The IP address of the proxy server

.PARAMETER proxyServerIpPort
    The IP port of the proxy server

.PARAMETER authToken
    The authentication token for connection

.PARAMETER correlationId
    The correlation ID for the connection

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $subscriptionId,
    [Parameter(Mandatory = $true)]
    [String]
    $resourceGroup,
    [Parameter(Mandatory = $true)]
    [String]
    $tenantId,
    [Parameter(Mandatory = $true)]
    [String]
    $azureRegion,
    [Parameter(Mandatory = $true)]
    [boolean]
    $useProxyServer,
    [Parameter(Mandatory = $false)]
    [String]
    $proxyServerIpAddress,
    [Parameter(Mandatory = $false)]
    [String]
    $proxyServerIpPort,
    [Parameter(Mandatory = $true)]
    [string]
    $authToken,
    [Parameter(Mandatory = $true)]
    [string]
    $correlationId
)

Set-StrictMode -Version 5.0

<#

.SYNOPSIS
Setup script runtime environment.

.DESCRIPTION
Setup script runtime environment.

#>

function setupScriptEnv() {
    Set-Variable -Name LogName -Option ReadOnly -Value "Microsoft-ServerManagementExperience" -Scope Script
    Set-Variable -Name LogSource -Option ReadOnly -Value "SMEScript" -Scope Script
    Set-Variable -Name ScriptName -Option ReadOnly -Value "Set-HybridManagement.ps1" -Scope Script
    Set-Variable -Name Machine -Option ReadOnly -Value "Machine" -Scope Script
    Set-Variable -Name HybridAgentFile -Option ReadOnly -Value "AzureConnectedMachineAgent.msi" -Scope Script
    Set-Variable -Name HybridAgentPackageLink -Option ReadOnly -Value "https://aka.ms/AzureConnectedMachineAgent" -Scope Script
    Set-Variable -Name HybridAgentExecutable -Option ReadOnly -Value "$env:ProgramFiles\AzureConnectedMachineAgent\azcmagent.exe" -Scope Script
    Set-Variable -Name HttpsProxy -Option ReadOnly -Value "https_proxy" -Scope Script
}

<#

.SYNOPSIS
Cleanup script runtime environment.

.DESCRIPTION
Cleanup script runtime environment.

#>

function cleanupScriptEnv() {
    Remove-Variable -Name LogName -Scope Script -Force
    Remove-Variable -Name LogSource -Scope Script -Force
    Remove-Variable -Name ScriptName -Scope Script -Force
    Remove-Variable -Name Machine -Scope Script -Force
    Remove-Variable -Name HybridAgentFile -Scope Script -Force
    Remove-Variable -Name HybridAgentPackageLink -Scope Script -Force
    Remove-Variable -Name HybridAgentExecutable -Scope Script -Force
    Remove-Variable -Name HttpsProxy -Scope Script -Force
}

<#

.SYNOPSIS
The main function.

.DESCRIPTION
Export the passed in virtual machine on this server.

#>

function main(
    [string]$subscriptionId,
    [string]$resourceGroup,
    [string]$tenantId,
    [string]$azureRegion,
    [boolean]$useProxyServer,
    [string]$proxyServerIpAddress,
    [string]$proxyServerIpPort,
    [string]$authToken,
    [string]$correlationId
) {
    $err = $null
    $args = @{}

    # Download the package
    Invoke-WebRequest -Uri $HybridAgentPackageLink -OutFile $HybridAgentFile -ErrorAction SilentlyContinue -ErrorVariable +err
    if ($err) {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]:Couldn't download the hybrid management package. Error: $err"  -ErrorAction SilentlyContinue

        Write-Error @($err)[0]
        return @()
    }

    # Install the package
    msiexec /i $HybridAgentFile /l*v installationlog.txt /qn | Out-String -ErrorAction SilentlyContinue -ErrorVariable +err
    if ($err) {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]:Error while installing the hybrid agent package. Error: $err"  -ErrorAction SilentlyContinue

        Write-Error @($err)[0]
        return @()
    }

    # Set the proxy environment variable. Note that authenticated proxies are not supported for Private Preview.
    if ($useProxyServer) {
        [System.Environment]::SetEnvironmentVariable($HttpsProxy, $proxyServerIpAddress+':'+$proxyServerIpPort, $Machine)
        $env:https_proxy = [System.Environment]::GetEnvironmentVariable($HttpsProxy, $Machine)
    }

    # Run connect command
    & $HybridAgentExecutable connect --resource-group $resourceGroup --tenant-id $tenantId --location $azureRegion `
                                     --subscription-id $subscriptionId --access-token $authToken --correlation-id $correlationId

    # Restart himds service
    Restart-Service -Name himds -ErrorAction SilentlyContinue -ErrorVariable +err
    if ($err) {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]:Couldn't restart the himds service. Error: $err"  -ErrorAction SilentlyContinue

        Write-Error @($err)[0]
        return $err
    }
}


###############################################################################
# Script execution starts here
###############################################################################
setupScriptEnv

try {
    Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

    return main $subscriptionId $resourceGroup $tenantId $azureRegion $useProxyServer $proxyServerIpAddress $proxyServerIpPort $authToken $correlationId

} finally {
    cleanupScriptEnv
}

}
## [END] Set-WACNSAzureHybridManagement ##
function Set-WACNSVMPovisioning {
<#

.SYNOPSIS
Prepare VM Provisioning

.DESCRIPTION
Prepare VM Provisioning

.ROLE
Administrators

#>

Param (
    [Parameter(Mandatory = $true)]
    [array]$disks
)

$output = @{ }

$requiredDriveLetters = $disks.driveLetter
$volumeLettersInUse = (Get-Volume | Sort-Object DriveLetter).DriveLetter

$output.Set_Item('restartNeeded', $false)
$output.Set_Item('pageFileLetterChanged', $false)
$output.Set_Item('pageFileLetterNew', $null)
$output.Set_Item('pageFileLetterOld', $null)
$output.Set_Item('pageFileDiskNumber', $null)
$output.Set_Item('cdDriveLetterChanged', $false)
$output.Set_Item('cdDriveLetterNew', $null)
$output.Set_Item('cdDriveLetterOld', $null)

$cdDriveLetterNeeded = $false
$cdDrive = Get-WmiObject -Class Win32_volume -Filter 'DriveType=5' | Microsoft.PowerShell.Utility\Select-Object -First 1
if ($cdDrive -ne $null) {
    $cdDriveLetter = $cdDrive.DriveLetter.split(':')[0]
    $output.Set_Item('cdDriveLetterOld', $cdDriveLetter)

    if ($requiredDriveLetters.Contains($cdDriveLetter)) {
        $cdDriveLetterNeeded = $true
    }
}

$pageFileLetterNeeded = $false
$pageFile = Get-WmiObject Win32_PageFileusage
if ($pageFile -ne $null) {
    $pagingDriveLetter = $pageFile.Name.split(':')[0]
    $output.Set_Item('pageFileLetterOld', $pagingDriveLetter)

    if ($requiredDriveLetters.Contains($pagingDriveLetter)) {
        $pageFileLetterNeeded = $true
    }
}

if ($cdDriveLetterNeeded -or $pageFileLetterNeeded) {
    $capitalCCharNumber = 67;
    $capitalZCharNumber = 90;

    for ($index = $capitalCCharNumber; $index -le $capitalZCharNumber; $index++) {
        $tempDriveLetter = [char]$index

        $willConflict = $requiredDriveLetters.Contains([string]$tempDriveLetter)
        $inUse = $volumeLettersInUse.Contains($tempDriveLetter)
        if (!$willConflict -and !$inUse) {
            if ($cdDriveLetterNeeded) {
                $output.Set_Item('cdDriveLetterNew', $tempDriveLetter)
                $cdDrive | Set-WmiInstance -Arguments @{DriveLetter = $tempDriveLetter + ':' } > $null
                $output.Set_Item('cdDriveLetterChanged', $true)
                $cdDriveLetterNeeded = $false
            }
            elseif ($pageFileLetterNeeded) {

                $computerObject = Get-WmiObject Win32_computersystem -EnableAllPrivileges
                $computerObject.AutomaticManagedPagefile = $false
                $computerObject.Put() > $null

                $currentPageFile = Get-WmiObject Win32_PageFilesetting
                $currentPageFile.delete() > $null

                $diskNumber = (Get-Partition -DriveLetter $pagingDriveLetter).DiskNumber

                $output.Set_Item('pageFileLetterNew', $tempDriveLetter)
                $output.Set_Item('pageFileDiskNumber', $diskNumber)
                $output.Set_Item('pageFileLetterChanged', $true)
                $output.Set_Item('restartNeeded', $true)
                $pageFileLetterNeeded = $false
            }

        }
        if (!$cdDriveLetterNeeded -and !$pageFileLetterNeeded) {
            break
        }
    }
}

# case where not enough drive letters available after iterating through C-Z
if ($cdDriveLetterNeeded -or $pageFileLetterNeeded) {
    $output.Set_Item('preProvisioningSucceeded', $false)
}
else {
    $output.Set_Item('preProvisioningSucceeded', $true)
}


Write-Output $output


}
## [END] Set-WACNSVMPovisioning ##
function Start-WACNSCimService {
<#

.SYNOPSIS
Start a service using CIM Win32_Service class.

.DESCRIPTION
Start a service using CIM Win32_Service class.

.ROLE
Administrators

#>

##SkipCheck=true##

Param(
[string]$Name
)

import-module CimCmdlets

$keyInstance = New-CimInstance -Namespace root/cimv2 -ClassName Win32_Service -Key @('Name') -Property @{Name=$Name;} -ClientOnly
Invoke-CimMethod $keyInstance -MethodName StartService

}
## [END] Start-WACNSCimService ##
function Start-WACNSVMProvisioning {
<#

.SYNOPSIS
Execute VM Provisioning

.DESCRIPTION
Execute VM Provisioning

.ROLE
Administrators

#>

Param (
    [Parameter(Mandatory = $true)]
    [bool] $partitionDisks,

    [Parameter(Mandatory = $true)]
    [array]$disks,

    [Parameter(Mandatory = $true)]
    [bool]$pageFileLetterChanged,

    [Parameter(Mandatory = $false)]
    [string]$pageFileLetterNew,

    [Parameter(Mandatory = $false)]
    [int]$pageFileDiskNumber,

    [Parameter(Mandatory = $true)]
    [bool]$systemDriveModified
)

$output = @{ }

$output.Set_Item('restartNeeded', $pageFileLetterChanged)

if ($pageFileLetterChanged) {
    Get-Partition -DiskNumber $pageFileDiskNumber | Set-Partition -NewDriveLetter $pageFileLetterNew
    $newPageFile = $pageFileLetterNew + ':\pagefile.sys'
    Set-WMIInstance -Class Win32_PageFileSetting -Arguments @{name = $newPageFile; InitialSize = 0; MaximumSize = 0 } > $null
}

if ($systemDriveModified) {
    $size = Get-PartitionSupportedSize -DriveLetter C
    Resize-Partition -DriveLetter C -Size $size.SizeMax > $null
}

if ($partitionDisks -eq $true) {
    $dataDisks = Get-Disk | Where-Object PartitionStyle -eq 'RAW' | Sort-Object Number
    for ($index = 0; $index -lt $dataDisks.Length; $index++) {
        Initialize-Disk  $dataDisks[$index].DiskNumber -PartitionStyle GPT -PassThru |
        New-Partition -Size $disks[$index].volumeSizeInBytes -DriveLetter $disks[$index].driveLetter |
        Format-Volume -FileSystem $disks[$index].fileSystem -NewFileSystemLabel $disks[$index].name -Confirm:$false -Force > $null;
    }
}

Write-Output $output

}
## [END] Start-WACNSVMProvisioning ##
function Suspend-WACNSCimService {
<#

.SYNOPSIS
Suspend a service using CIM Win32_Service class.

.DESCRIPTION
Suspend a service using CIM Win32_Service class.

.ROLE
Administrators

#>

##SkipCheck=true##

Param(
[string]$Name
)

import-module CimCmdlets

$keyInstance = New-CimInstance -Namespace root/cimv2 -ClassName Win32_Service -Key @('Name') -Property @{Name=$Name;} -ClientOnly
Invoke-CimMethod $keyInstance -MethodName PauseService

}
## [END] Suspend-WACNSCimService ##

# SIG # Begin signature block
# MIIoOQYJKoZIhvcNAQcCoIIoKjCCKCYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCkeHH9X+PAiMOr
# 5taEVc4b+XcdE3EjC7cmsJVGaB4nSaCCDYUwggYDMIID66ADAgECAhMzAAADTU6R
# phoosHiPAAAAAANNMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMwMzE2MTg0MzI4WhcNMjQwMzE0MTg0MzI4WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDUKPcKGVa6cboGQU03ONbUKyl4WpH6Q2Xo9cP3RhXTOa6C6THltd2RfnjlUQG+
# Mwoy93iGmGKEMF/jyO2XdiwMP427j90C/PMY/d5vY31sx+udtbif7GCJ7jJ1vLzd
# j28zV4r0FGG6yEv+tUNelTIsFmmSb0FUiJtU4r5sfCThvg8dI/F9Hh6xMZoVti+k
# bVla+hlG8bf4s00VTw4uAZhjGTFCYFRytKJ3/mteg2qnwvHDOgV7QSdV5dWdd0+x
# zcuG0qgd3oCCAjH8ZmjmowkHUe4dUmbcZfXsgWlOfc6DG7JS+DeJak1DvabamYqH
# g1AUeZ0+skpkwrKwXTFwBRltAgMBAAGjggGCMIIBfjAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUId2Img2Sp05U6XI04jli2KohL+8w
# VAYDVR0RBE0wS6RJMEcxLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJh
# dGlvbnMgTGltaXRlZDEWMBQGA1UEBRMNMjMwMDEyKzUwMDUxNzAfBgNVHSMEGDAW
# gBRIbmTlUAXTgqoXNzcitW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8v
# d3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIw
# MTEtMDctMDguY3JsMGEGCCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDEx
# XzIwMTEtMDctMDguY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIB
# ACMET8WuzLrDwexuTUZe9v2xrW8WGUPRQVmyJ1b/BzKYBZ5aU4Qvh5LzZe9jOExD
# YUlKb/Y73lqIIfUcEO/6W3b+7t1P9m9M1xPrZv5cfnSCguooPDq4rQe/iCdNDwHT
# 6XYW6yetxTJMOo4tUDbSS0YiZr7Mab2wkjgNFa0jRFheS9daTS1oJ/z5bNlGinxq
# 2v8azSP/GcH/t8eTrHQfcax3WbPELoGHIbryrSUaOCphsnCNUqUN5FbEMlat5MuY
# 94rGMJnq1IEd6S8ngK6C8E9SWpGEO3NDa0NlAViorpGfI0NYIbdynyOB846aWAjN
# fgThIcdzdWFvAl/6ktWXLETn8u/lYQyWGmul3yz+w06puIPD9p4KPiWBkCesKDHv
# XLrT3BbLZ8dKqSOV8DtzLFAfc9qAsNiG8EoathluJBsbyFbpebadKlErFidAX8KE
# usk8htHqiSkNxydamL/tKfx3V/vDAoQE59ysv4r3pE+zdyfMairvkFNNw7cPn1kH
# Gcww9dFSY2QwAxhMzmoM0G+M+YvBnBu5wjfxNrMRilRbxM6Cj9hKFh0YTwba6M7z
# ntHHpX3d+nabjFm/TnMRROOgIXJzYbzKKaO2g1kWeyG2QtvIR147zlrbQD4X10Ab
# rRg9CpwW7xYxywezj+iNAc+QmFzR94dzJkEPUSCJPsTFMIIHejCCBWKgAwIBAgIK
# YQ6Q0gAAAAAAAzANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlm
# aWNhdGUgQXV0aG9yaXR5IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEw
# OTA5WjB+MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYD
# VQQDEx9NaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+la
# UKq4BjgaBEm6f8MMHt03a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc
# 6Whe0t+bU7IKLMOv2akrrnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4D
# dato88tt8zpcoRb0RrrgOGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+
# lD3v++MrWhAfTVYoonpy4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nk
# kDstrjNYxbc+/jLTswM9sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6
# A4aN91/w0FK/jJSHvMAhdCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmd
# X4jiJV3TIUs+UsS1Vz8kA/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL
# 5zmhD+kjSbwYuER8ReTBw3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zd
# sGbiwZeBe+3W7UvnSSmnEyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3
# T8HhhUSJxAlMxdSlQy90lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS
# 4NaIjAsCAwEAAaOCAe0wggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRI
# bmTlUAXTgqoXNzcitW2oynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTAL
# BgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBD
# uRQFTuHqp8cx0SOJNDBaBgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jv
# c29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFf
# MDNfMjIuY3JsMF4GCCsGAQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFf
# MDNfMjIuY3J0MIGfBgNVHSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEF
# BQcCARYzaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1h
# cnljcHMuaHRtMEAGCCsGAQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkA
# YwB5AF8AcwB0AGEAdABlAG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn
# 8oalmOBUeRou09h0ZyKbC5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7
# v0epo/Np22O/IjWll11lhJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0b
# pdS1HXeUOeLpZMlEPXh6I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/
# KmtYSWMfCWluWpiW5IP0wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvy
# CInWH8MyGOLwxS3OW560STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBp
# mLJZiWhub6e3dMNABQamASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJi
# hsMdYzaXht/a8/jyFqGaJ+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYb
# BL7fQccOKO7eZS/sl/ahXJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbS
# oqKfenoi+kiVH6v7RyOA9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sL
# gOppO6/8MO0ETI7f33VtY5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtX
# cVZOSEXAQsmbdlsKgEhr/Xmfwb1tbWrJUnMTDXpQzTGCGgowghoGAgEBMIGVMH4x
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01p
# Y3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTECEzMAAANNTpGmGiiweI8AAAAA
# A00wDQYJYIZIAWUDBAIBBQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIOGr
# PDwCjmg2iNqsU0e/MqLnsYt8zcPMxwUpGM93zQJQMEIGCisGAQQBgjcCAQwxNDAy
# oBSAEgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20wDQYJKoZIhvcNAQEBBQAEggEACl0Wea2z0dUxw1HYb0TMDExqkEc3TsDIiMmb
# xS4g95XO/m+Kgw+g5CgryCCb4MzRUckTcvvQWflIxAJ42dD3RtzM5tQcJWlfX//n
# JUPQJXRfmfaUuIaMmh3V9W+6+YHD6qfcHCZ9TJi9KoKujPWls/EzxLhC6tCNXCiI
# 5mVHiZpYiKOhe4SRwaFHs7Nrd7UnMIOEWR9UL3VsXmxSnXJDPQD3aHLZ7tfzVJ/D
# 1CDU1y9mSQPCEVakjybOGKG1j+OuQ4eWQSJVw45C59SQLcpPO5nhl5aUewCdixhv
# sVljpozpjkG0wz86Aqr7mlzvPeOS1Qbbt+0AccLdqdIakNfbsqGCF5QwgheQBgor
# BgEEAYI3AwMBMYIXgDCCF3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZI
# AWUDBAIBBQAwggFSBgsqhkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGE
# WQoDATAxMA0GCWCGSAFlAwQCAQUABCB8COjt455bPocD7YH1+qAX5zM/xM6Eu+uU
# +G+m7YoarAIGZVbJDY0GGBMyMDIzMTIwNzA1MTA1Ny40MzNaMASAAgH0oIHRpIHO
# MIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQL
# ExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxk
# IFRTUyBFU046QTAwMC0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1l
# LVN0YW1wIFNlcnZpY2WgghHqMIIHIDCCBQigAwIBAgITMwAAAdB3CKrvoxfG3QAB
# AAAB0DANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAx
# MDAeFw0yMzA1MjUxOTEyMTRaFw0yNDAyMDExOTEyMTRaMIHLMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1l
# cmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046QTAwMC0w
# NUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Uw
# ggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDfMlfn35fvM0XAUSmI5qiG
# 0UxPi25HkSyBgzk3zpYO311d1OEEFz0QpAK23s1dJFrjB5gD+SMw5z6EwxC4CrXU
# 9KaQ4WNHqHrhWftpgo3MkJex9frmO9MldUfjUG56sIW6YVF6YjX+9rT1JDdCDHbo
# 5nZiasMigGKawGb2HqD7/kjRR67RvVh7Q4natAVu46Zf5MLviR0xN5cNG20xwBwg
# ttaYEk5XlULaBH5OnXz2eWoIx+SjDO7Bt5BuABWY8SvmRQfByT2cppEzTjt/fs0x
# p4B1cAHVDwlGwZuv9Rfc3nddxgFrKA8MWHbJF0+aWUUYIBR8Fy2guFVHoHeOze7I
# sbyvRrax//83gYqo8c5Z/1/u7kjLcTgipiyZ8XERsLEECJ5ox1BBLY6AjmbgAzDd
# Nl2Leej+qIbdBr/SUvKEC+Xw4xjFMOTUVWKWemt2khwndUfBNR7Nzu1z9L0Wv7TA
# Y/v+v6pNhAeohPMCFJc+ak6uMD8TKSzWFjw5aADkmD9mGuC86yvSKkII4MayzoUd
# seT0nfk8Y0fPjtdw2Wnejl6zLHuYXwcDau2O1DMuoiedNVjTF37UEmYT+oxC/OFX
# UGPDEQt9tzgbR9g8HLtUfEeWOsOED5xgb5rwyfvIss7H/cdHFcIiIczzQgYnsLyE
# GepoZDkKhSMR5eCB6Kcv/QIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFDPhAYWS0oA+
# lOtITfjJtyl0knRRMB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8G
# A1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMv
# Y3JsL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBs
# BggrBgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0
# LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUy
# MDIwMTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUH
# AwgwDgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQCXh+ckCkZaA06S
# NW+qxtS9gHQp4x7G+gdikngKItEr8otkXIrmWPYrarRWBlY91lqGiilHyIlZ3iNB
# UbaNEmaKAGMZ5YcS7IZUKPaq1jU0msyl+8og0t9C/Z26+atx3vshHrFQuSgwTHZV
# pzv7k8CYnBYoxdhI1uGhqH595mqLvtMsxEN/1so7U+b3U6LCry5uwwcz5+j8Oj0G
# UX3b+iZg+As0xTN6T0Qa8BNec/LwcyqYNEaMkW2VAKrmhvWH8OCDTcXgONnnABQH
# BfXK/fLAbHFGS1XNOtr62/iaHBGAkrCGl6Bi8Pfws6fs+w+sE9r3hX9Vg0gsRMoH
# RuMaiXsrGmGsuYnLn3AwTguMatw9R8U5vJtWSlu1CFO5P0LEvQQiMZ12sQSsQAkN
# DTs9rTjVNjjIUgoZ6XPMxlcPIDcjxw8bfeb4y4wAxM2RRoWcxpkx+6IIf2L+b7gL
# HtBxXCWJ5bMW7WwUC2LltburUwBv0SgjpDtbEqw/uDgWBerCT+Zty3Nc967iGaQj
# yYQH6H/h9Xc8smm2n6VjySRx2swnW3hr6Qx63U/xY9HL6FNhrGiFED7ZRKrnwvvX
# vMVQUIEkB7GUEeN6heY8gHLt0jLV3yzDiQA8R8p5YGgGAVt9MEwgAJNY1iHvH/8v
# zhJSZFNkH8svRztO/i3TvKrjb8ZxwjCCB3EwggVZoAMCAQICEzMAAAAVxedrngKb
# SZkAAAAAABUwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmlj
# YXRlIEF1dGhvcml0eSAyMDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIy
# NVowfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcT
# B1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UE
# AxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qlsTnXI
# yjVX9gF/bErg4r25PhdgM/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLAEBjo
# YH1qUoNEt6aORmsHFPPFdvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrEqv1y
# aa8dq6z2Nr41JmTamDu6GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v
# 3byNpOORj7I5LFGc6XBpDco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1oO5pG
# ve2krnopN6zL64NF50ZuyjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viS
# kR4dPf0gz3N9QZpGdc3EXzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2TPYr
# bqgSUei/BQOj0XOmTTd0lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlM
# jgK8QmguEOqEUUbi0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJNmSL
# W6CmgyFdXzB0kZSU2LlQ+QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AF
# emzFER1y7435UsSFF5PAPBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+auIu
# rQIDAQABo4IB3TCCAdkwEgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIE
# FgQUKqdS/mTEmr6CkTxGNSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl0mWn
# G1M1GelyMFwGA1UdIARVMFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUHAgEW
# M2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5
# Lmh0bTATBgNVHSUEDDAKBggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMAdQBi
# AEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV
# 9lbLj+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3Js
# Lm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAx
# MC0wNi0yMy5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8v
# d3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2
# LTIzLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL/Klv
# 6lwUtj5OR2R4sQaTlz0xM7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZn
# OlNN3Zi6th542DYunKmCVgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5tggz1
# bSNU5HhTdSRXud2f8449xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfgQJY4
# rPf5KYnDvBewVIVCs/wMnosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU
# 6ZGyqVvfSaN0DLzskYDSPeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCrdTDF
# NLB62FD+CljdQDzHVG2dY3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/
# HltEAY5aGZFrDZ+kKNxnGSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2tVdU
# CbFpAUR+fKFhbHP+CrvsQWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKi
# excdFYmNcP7ntdAoGokLjzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9JZTm
# dHRbatGePu1+oDEzfbzL6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZq
# ELQdVTNYs6FwZvKhggNNMIICNQIBATCB+aGB0aSBzjCByzELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJp
# Y2EgT3BlcmF0aW9uczEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOkEwMDAtMDVF
# MC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMK
# AQEwBwYFKw4DAhoDFQC8t8hT8KKUX91lU5FqRP9Cfu9MiaCBgzCBgKR+MHwxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jv
# c29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6Rui1jAi
# GA8yMDIzMTIwNzAxNDgwNloYDzIwMjMxMjA4MDE0ODA2WjB0MDoGCisGAQQBhFkK
# BAExLDAqMAoCBQDpG6LWAgEAMAcCAQACAh4rMAcCAQACAhNvMAoCBQDpHPRWAgEA
# MDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAI
# AgEAAgMBhqAwDQYJKoZIhvcNAQELBQADggEBAE3W7+J48J388zug2M7n9821tMlm
# RDyzEiYulAAPYTv/Yvha83QyrNy8zNFQS+xoyPVLEdBKHO4O+WpGAHp5tpIDH8FD
# SCXE5AMZzY3yDRn9eWD1QvqgjcVpGYTEfQ/bvySPeA5TgTCsf9pGr4P2z0sPbNBP
# RStw4KLgbahGzq1zEAoiUmUnOa1Q9rlhEXjLPFWuOsDU5XCv/zHtDgs54Q14nPJI
# euO0LHFbeSgUCjXJVNpHSxUZazepPojq27vQpLKY88+xzelXCpJ+y8jRZIBxMFtC
# lSW+yTPeUDFZMRkvrPPPsRSKSETTd46Er8o6PBRx5jgWU9m5sLk23rBDJmAxggQN
# MIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAdB3
# CKrvoxfG3QABAAAB0DANBglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0G
# CyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCBPZC+uj361IJDfa6wXQGjCLZOE
# KC1KTW6LH/AXU2ruxzCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIAiVQAZf
# tNP/Md1E2Yw+fBXa9w6fjmTZ5WAerrTSPwnXMIGYMIGApH4wfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTACEzMAAAHQdwiq76MXxt0AAQAAAdAwIgQg1sweGpbV
# 3KSV3i3+HR8l4A0emIkE3P5XWHD5mcsKU8gwDQYJKoZIhvcNAQELBQAEggIAc341
# JQJQ4K3PWn9G+TivamqjdwIuPrBIxDTuRlFZU7ciqotJ0EjywuYZLqAO6zaVT9iu
# 5GKWW4tc/N5yBLhcZgLXh0RxEXQvJ5v/j9HNL+K2I/NuTLfKeFUtGsaCWS5jE/4u
# uzwliZmAiuOz5JBkmgVl7gQbHxzdPu/MQRk8RFGcB1GF5zA7QBVyuq2/+ad6qrVC
# O8ePPpt+u3AoFO4q7BRPlFsRLQOv2g90TqnUaO4LQ1xREKM3r5y7p6m5m2Ne/K0F
# 87wOfL35eAIYfUjhxmGI59KxjFYXZ6HFsQHuUt3mnQ7N8TXp/VUCMDRlEFh34E9J
# NSNNI6xzdVJfJ8DcVoMvNXBa7om3huhSSq5ujZ2YarwreMx96ry9SXK6aMzyKcwd
# Tk86cO8VPthjzhhuxXrwXGxBXfylO3YzkBfMeaMfx9eBH6/FdM0RT+ftc+M9z0YH
# WqFFcZ7VZUB2rcKHOd/EvY5xEFwr8Ezt5Ho9Uc0cUoRquyip1X2mAYHFz31BFffO
# 2d5psFzJQdN/+pQeSoY1zue65YrEQZmSpMBCts6ZF1fN2u2xyn9WcGjOwk+Dah+X
# LhqnYL8pYTMMj5c4ninW7PMgoG0rWGoCvpe7WIZZxp2uhkfRvVqq2cwac2KMbD+5
# srTm8zOBFnVthAQLQSTmO2hjeFYH722OKCz+CsQ=
# SIG # End signature block

Function Query-SoftwareInstalled
{
[CmdletBinding (SupportsShouldProcess = $True)]
Param
(
    [Parameter (Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True,
        HelpMessage="Input a domain OU structure to query for software installed.`r`nE.g.- `"OU=1stOUName,OU=2ndOUName,DC=LowLevelDomain,DC=MidLevelDomain,DC=TopLevelDomain`"")]
    [Alias('OU')]
    [string]$OUStructure,

    [Parameter (Mandatory=$True,
    ValueFromPipeline=$False,
    ValueFromPipelineByPropertyName=$False,
        HelpMessage="Input the software name that is to be queried.`r`n(Be sure it matches the software's name listed in the registry.)")]
    [Alias('Install')]
    [string[]]$Software,

    [Parameter (Mandatory=$False,
    ValueFromPipeline=$False,
    ValueFromPipelineByPropertyName=$False,
        HelpMessage="Input the number of days before a computer account is considered 'Inactive'.`r`nE.g.- `"30`"")]
    [Alias('Days')]
    [int32]$InactivityThreshold = "30",

    [Parameter (Mandatory=$False,
    ValueFromPipeline=$False,
    ValueFromPipelineByPropertyName=$False,
        HelpMessage="Input a directory to output the CSV file.")]
    [Alias('Folder')]
    [string]$OutputPath = "$env:USERPROFILE\Desktop\$Software-Machines"
)
$date = Get-Date -Format MMM-dd-yyyy
$time = [DateTime]::Now
$Computers = Get-ADComputer -Filter {LastLogonTimeStamp -lt $time} -SearchBase "$OUStructure" -Properties 'Name','OperatingSystem','CanonicalName','LastLogonTimeStamp'
If ($OutputPath)
{
    If (-not (Test-Path -Path $OutputPath))
    {
    New-Item -Path "$OutputPath" -ItemType Directory
    }
}
ForEach ($Computer in $Computers)
{
$subkeyarray = $null
$NetAdapterError = $null
$nocomp = $null
$comp = $null
$Name = $null
$IPAddress = $null
$OS = $null
$CanonicalName = $null
$DisplayName = $null
$DisplayVersion = $null
$InstallLocation = $null
$Publisher = $null
$NetConfig = $null
$MAC = $null
$IPEnabled = $null
$DNSServers = $null
$LogonTime = [DateTime]::FromFileTime($Computer.LastLogonTimeStamp)
    If ($LogonTime -gt (Get-Date).AddDays(-($InactivityThreshold)))
    {
    $Name = $($Computer.Name)
    $NetConfig = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $($Computer.Name) -ErrorAction SilentlyContinue -ErrorVariable NetAdapterError | Where {$_.IPEnabled -eq $true}
        If ($NetAdapterError -like "*The RPC server is unavailable*")
        {
        $IPAddress = "ERROR: Remote connection to $($Computer.Name)`'s network adapter failed."
        }
        Else
        {
            ForEach ($AdapterItem in $NetConfig)
            {
            $MAC = $AdapterItem.MACAddress
            $IPAddress = $AdapterItem.IPAddress | Where {$_ -like "172.*"}
            $IPEnabled = $AdapterItem.IPEnabled
            $DNSServers = $AdapterItem.DNSServerSearchOrder
            }
        }
    $OS = $($Computer.OperatingSystem)
    $CanonicalName = $($Computer.CanonicalName)
    #Define the variable to hold the location of Currently Installed Programs
    $UninstallKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
    #Create an instance of the Registry Object and open the HKLM base key
    $subkeyarray = @()
    $PrevErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "SilentlyContinue"
    Write-Error -Message "Test - Disregard"
    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$($Computer.Name))
    $ErrorActionPreference = $PrevErrorActionPreference
        If ($error[0] -like "*Exception calling `"OpenRemoteBaseKey`"*" -and $error[0] -like "*`"The network path was not found.*")
        {
        $DisplayName = "No Network connection to $($Computer.Name)!"
        $DisplayVersion = "No Network connection to $($Computer.Name)!"
        $InstallLocation = "No Network connection to $($Computer.Name)!"
        $Publisher = "No Network connection to $($Computer.Name)!"
        $obj = New-Object PSObject
        $obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value "No Network connection to $($Computer.Name)!"
        $obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value "No Network connection to $($Computer.Name)!"
        $obj | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value "No Network connection to $($Computer.Name)!"
        $obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value "No Network connection to $($Computer.Name)!"
        $subkeyarray += $obj
        $obj = $null
        }
        Else
        {
        #Drill down into the Uninstall key using the OpenSubKey Method
        $regkey = $reg.OpenSubKey($UninstallKey)
        #Retrieve an array of strings that contain all the subkey names
        $subkeys = $regkey.GetSubKeyNames() 
        #Open each Subkey and use GetValue Method to return the required values for each
            ForEach ($key in $subkeys)
            {
            $thisKey = $UninstallKey + "\\" + $key
            $thisSubKey = $reg.OpenSubKey($thisKey)
                If ($($thisSubKey.GetValue("DisplayName")) -like "*$Software*")
                {
                $obj = New-Object PSObject
                $obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
                $obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
                $obj | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value $($thisSubKey.GetValue("InstallLocation"))
                $obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $($thisSubKey.GetValue("Publisher"))
                $subkeyarray += $obj
                $obj = $null
                }
            }
            If ($subkeyarray.DisplayName -notlike "*No Network connection*" -and $subkeyarray.DisplayName -notlike "*$Software*")
            {
            $obj = New-Object PSObject
            $obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value "$Software not installed!"
            $obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value "$Software not installed!"
            $obj | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value "$Software not installed!"
            $obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value "$Software not installed!"
            $subkeyarray += $obj
            $obj = $null
            }
        $DisplayName = [string]::Concat($subkeyarray.DisplayName)
        $DisplayVersion = [string]::Concat($subkeyarray.DisplayVersion)
        $InstallLocation = [string]::Concat($subkeyarray.InstallLocation)
        $Publisher = [string]::Concat($subkeyarray.Publisher)
        }
    $comp = New-Object PSObject
    $comp | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $Name
    $comp | Add-Member -MemberType NoteProperty -Name "IP_Address" -Value $IPAddress
    $comp | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS
    $comp | Add-Member -MemberType NoteProperty -Name "OUStructure" -Value $CanonicalName
    $comp | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $DisplayName
    $comp | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $DisplayVersion
    $comp | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value $InstallLocation
    $comp | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $Publisher
    $comp | Export-Csv -Path $OutputPath\$Software-Machines_$date.csv -Encoding ascii -Append -Force
    }
    Else
    {
    $nocomp = New-Object PSObject
    $nocomp | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value "$($Computer.Name) has not contacted AD in over $InactivityThreshold days"
    $nocomp | Add-Member -MemberType NoteProperty -Name "IP_Address" -Value "ERROR"
    $nocomp | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value "ERROR"
    $nocomp | Add-Member -MemberType NoteProperty -Name "OUStructure" -Value "ERROR"
    $nocomp | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value "ERROR"
    $nocomp | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value "ERROR"
    $nocomp | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value "ERROR"
    $nocomp | Add-Member -MemberType NoteProperty -Name "Publisher" -Value "ERROR"
    $nocomp | Export-Csv -Path $OutputPath\$Software-Machines_$date.csv -Encoding ascii -Append -Force
    }
}
}
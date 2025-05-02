<#
 NAME  : Get-VMInventory.ps1
 AUTHOR: Jake Dennis
 DATE  : 3/18/2019
 DESCRIPTION
	This script will scan the provided vCenter(s) with the option to filter VMs and output basic information to a CSV file.
 EXAMPLE
	Get-VMInventory -vCtrs 'vctr1','vctr2' -FilePath 'C:\temp\2003Servers.csv' -OSFilter '*Microsoft Windows Server 2003*'
 LINK
	https://github.com/JakeDennis/
#>
function Get-VMInventory{
    [cmdletbinding()]
    Param(
        [parameter(Position=2)]
        $OSFilter,
        [parameter(Mandatory=$true, Position=1)]
        $FilePath,
        [parameter(Mandatory=$true, Position=0)]
        $vCtrs
        ) 
    
    #$Creds = Get-Credential -Message "Enter credentials to connect to $($vCtr)." -UserName "yellowcorpnt\sa_jdennis"
    
    #Loop to scan each vCenter for respective VMs
    foreach($vCtr in $vCtrs){
        
        try{
        Connect-VIServer -Server $vCtr -User "administrator@vsphere.local" -Password "J@bb3rJ4w" -ErrorAction Continue
        }
        catch{
        Write-Host "There was an error connecting to $($vCtr)."
        }
        
        #Gather VMs from vCenter
        $VMs = Get-VM
        $Count = 0
        $Servers = @()
        foreach($VM in $VMs){
            try{
            $VMGuest = Get-VMGuest -VM $VM.Name
            }
            catch{
            Write-Host "There was an error processing $($VM)."
            }

            #Create table for VM Data
            if($VMGuest.OSFullName -like $OSFilter){
                $VMHost = Get-VMHost -VM $VM.Name
                $Server = New-Object PSCustomObject
                [string]$IPAddress = $VMGuest.IPAddress
                [int]$Size = $VM.ProvisionedSpaceGB

                    #Data set
                    $Server | Add-Member -MemberType NoteProperty -Name ServerName -Value $VM.Name
                    $Server | Add-Member -MemberType NoteProperty -Name IPAddress -Value $IPAddress
                    $Server | Add-Member -MemberType NoteProperty -Name FQDN -Value $VMGuest.Hostname
                    $Server | Add-Member -MemberType NoteProperty -Name Host -Value $VMHost.Name
                    $Server | Add-Member -MemberType NoteProperty -Name Cluster -Value $VMHost.Parent
                    $Server | Add-Member -MemberType NoteProperty -Name vCenter -Value $vCtr
                    $Server | Add-Member -MemberType NoteProperty -Name State -Value $VMGuest.State
                    $Server | Add-Member -MemberType NoteProperty -Name OperatingSystem -Value $VMGuest.OSFullName
                    $Server | Add-Member -MemberType NoteProperty -Name CPUs -Value $VM.NumCpu
                    $Server | Add-Member -MemberType NoteProperty -Name 'Memory(GB)' -Value $VM.MemoryGB
                    $Server | Add-Member -MemberType NoteProperty -Name 'Size(GB)' -Value $Size
                    $Server | Add-Member -MemberType NoteProperty -Name Network -Value $VMGuest.Nics.NetworkName
                    $Server | Add-Member -MemberType NoteProperty -Name MACAddress -Value $VMGuest.Nics.MacAddress
                    $Server 
                    $Servers += $Server       
                    }

            #Progress Bar
            $Count++
            $Parameters = @{Activity = "Processing: "
                            Status = "$($vCtr) - $($Count) of $($VMs.Count) VMs."
                            CurrentOperation = $VM
                            PercentComplete = (($Count/$VMs.Count)*100)
                            }
            Write-Progress @Parameters
            }
    }
    $Servers | Export-Csv -NoTypeInformation -Path $FilePath    
    Write-Progress -Activity "Processing: " -Completed
} 

    
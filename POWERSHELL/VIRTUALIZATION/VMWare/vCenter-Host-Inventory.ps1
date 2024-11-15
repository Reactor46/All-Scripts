﻿<# 
*******************************************************************************************************************
Authored Date:    March 2018
Original Author:  Graham Jensen
*******************************************************************************************************************
.SYNOPSIS
    Get Host listing from a vCenter

.DESCRIPTION
   Generates an Excel Worksheet with all the Hosts Managed by a particular vCenter.  Includes the gathering 
   of Annotations, and gathers available information from the hosts, such as Manufacture, Model, Serial #, 
   BIOS info, ESXi Version and other into.

   Prompted inputs:  Credentials, vCenterName

   Outputs:          
            $USERPROFILE$\Documents\vCenterHostListings\$VMHost-HostList.xlsx

*******************************************************************************************************************  
.NOTES
Prerequisites:

    #1  This script uses the VMware modules installed by the installation of VMware PowerCLI
        ENSURE that VMware PowerCLI has been installed.  

===================================================================================================================
Update Log:   Please use this section to document changes made to this script
===================================================================================================================
-----------------------------------------------------------------------------
Update <Date>
   Author:    <Name>
   Description of Change:
      <Description>
-----------------------------------------------------------------------------
*******************************************************************************************************************
#>

# -----------------------
# Define Global Variables
# -----------------------
$Global:Folder = $env:USERPROFILE+"\Documents\vCenter" 
$Global:MasterList = $null
$Global:HostList = @()
$Global:VCName = $null
$Global:Creds = $null



#*****************
# Get VC from User
#*****************
Function Get-VCenter {
    [CmdletBinding()]
    Param()
    #Prompt User for vCenter
    Write-Host "Enter the FQHN of the vCenter to Get Hosting Listing From: " -ForegroundColor "Yellow" -NoNewline
    $Global:VCName = Read-Host 
}
#*******************
# EndFunction Get-VC
#*******************

#*************************************************
# Check for Folder Structure if not present create
#*************************************************
Function Verify-Folders {
    [CmdletBinding()]
    Param()
    "Building Local folder structure" 
    If (!(Test-Path $Global:Folder)) {
        New-Item $Global:Folder -type Directory
  <#      New-Item "$Global:WorkFolder\Annotations" -type Directory
        New-Item "$Global:WorkFolder\IPConfig" -type Directory
        New-Item "$Global:WorkFolder\ShareInfo" -type Directory
        New-Item "$Global:WorkFolder\VMInfo" -type Directory
        New-Item "$Global:WorkFolder\PrinterInfo" -type Directory
    #>
        }
    "Folder Structure built" 
}
#***************************
# EndFunction Verify-Folders
#***************************

#*******************
# Connect to vCenter
#*******************
Function Connect-VC {
    [CmdletBinding()]
    Param()
    "Connecting to $Global:VCName"
    Connect-VIServer $Global:VCName -Credential $Global:Creds -WarningAction SilentlyContinue
}
#***********************
# EndFunction Connect-VC
#***********************

#*******************
# Disconnect vCenter
#*******************
Function Disconnect-VC {
    [CmdletBinding()]
    Param()
    "Disconnecting $Global:VCName"
    Disconnect-VIServer -Server $Global:VCName -Confirm:$false
}
#**************************
# EndFunction Disconnect-VC
#**************************


#*********************
# Clean Up after Run
#*********************
Function Clean-Up {
    [CmdletBinding()]
    Param()
    $Global:Folder = $null
    $Global:HostList = $null
    $Global:VCName = $null
    $Global:Creds = $null
}
#*********************
# EndFunction Clean-Up
#*********************

#***********************
# Function Get-IODevices
#***********************
Function Get-IODevices{
    [CmdletBinding()]
    Param($vmHost)
    $AllInfo = @()

    $esxcli = Get-ESXCLI -VMHost $vmHost -V2
    $niclist = $esxcli.network.nic.list.Invoke() 

    Foreach ($nic in $niclist){
        $Info = "" | Select Description, Driver, DriverVersion, FirmwareVersion, VibVersion
     
        $vmnicDetail = $esxcli.network.nic.get.Invoke(@{nicname = $nic.Name})
        $Info.Description = $nic.Description
        $Info.Driver = $vmnicDetail.DriverInfo.Driver
        $Info.DriverVersion = $vmnicDetail.DriverInfo.Version
        $Info.FirmwareVersion = $vmnicDetail.DriverInfo.FirmwareVersion
      
        # Get driver vib package version
        Try{
            $driverVib = $esxcli.software.vib.get.Invoke(@{vibname = "net-"+$vmnicDetail.DriverInfo.Driver})
            }
        Catch{
            $driverVib = $esxcli.software.vib.get.Invoke(@{vibname = $vmnicDetail.DriverInfo.Driver})
            }
        $Info.VibVersion = $driverVib.Version
        $AllInfo += $Info
   }
   Return $AllInfo
}
#***************************
# End-Function Get-IODevices
#***************************

#**********************
# Function Get-HostList
#**********************
Function Get-HostList {
    [CmdletBinding()]
    Param()
    "Generating Host View from $Global:VCname"
    "This may take a few minutes"
    $Count = 1
    #$Global:MasterList = Get-View -ViewType HostSystem -Filter @{"Name" = "itspaublesx101.cihs.gov.on.ca"}
    $Global:MasterList = Get-View -ViewType HostSystem
    "----------------------------------------------------------"
    "Generating Host Details from Host View"
    ForEach ($vmview in $Global:MasterList){
        Write-Progress -Id 0 -Activity 'Generating Host Details from Host View' -Status "Processing $($count) of $($Global:MasterList.count)" -CurrentOperation $_.Name -PercentComplete (($count/$Global:MasterList.count) * 100)
        $vmhost=New-Object PsObject

        $IOInfo = Get-Iodevices $vmview.name

        $vmhost | Add-Member -MemberType NoteProperty -Name FQDN -Value $vmview.Name
        $vmhost | Add-Member -MemberType NoteProperty -Name ShortName -Value $vmview.Name.Split(".")[0]
        $vmhost | Add-Member -MemberType NoteProperty -Name State -Value $vmview.Runtime.ConnectionState
        $vmhost | Add-Member -MemberType NoteProperty -Name Vendor -Value $vmview.Hardware.systemInfo.Vendor
        $vmhost | Add-Member -MemberType NoteProperty -Name Model -Value $vmview.Hardware.systemInfo.Model
        $vmhost | Add-Member -MemberType NoteProperty -Name Serial# -Value ($vmview.Hardware.SystemInfo.OtherIdentifyingInfo | Where {$_.IdentifierType.Key -eq "ServiceTag"}).IdentifierValue
        $vmhost | Add-Member -MemberType NoteProperty -Name BiosVersion -Value $vmview.Hardware.BiosInfo.BiosVersion
        $vmhost | Add-Member -MemberType NoteProperty -Name BiosReleaseDate -Value $vmview.Hardware.BiosInfo.ReleaseDate
        $vmhost | Add-Member -MemberType NoteProperty -Name RACFirmware -Value ($vmview.runtime.healthsystemruntime.systemhealthinfo.numericsensorinfo | where {$_.Name -Like "*BMC Firmware*"}).Name.Split(" ")[($_.count)-1]
        $SmartArray = ($vmview.runtime.healthsystemruntime.systemhealthinfo.numericsensorinfo | where {$_.Name -Like "*Smart Array*"}).Name
        if ($SmartArray.count -le 0){
            $vmhost | Add-Member -MemberType NoteProperty -Name SmartArrayFirmware -Value "N/A"
            }
        if ($SmartArray.count -gt 1){
            $vmhost | Add-Member -MemberType NoteProperty -Name SmartArrayFirmware -Value $SmartArray[0].split(" ")[($_.count)-1]
            }
            Else{
                $vmhost | Add-Member -MemberType NoteProperty -Name SmartArrayFirmware -Value $SmartArray.split(" ")[($_.count)-1]
            }
        if ($IOInfo.count -eq 1){
            $vmhost | Add-Member -MemberType NoteProperty -Name NetworkDescription -Value $IOInfo.Description
            $vmhost | Add-Member -MemberType NoteProperty -Name NetworkDrv -Value $IOInfo.Driver
            $vmhost | Add-Member -MemberType NoteProperty -Name NetworkDrvVer -Value $IOInfo.DriverVersion
            $vmhost | Add-Member -MemberType NoteProperty -Name NetworkFirmware -Value $IOInfo.FirmwareVersion
            $vmhost | Add-Member -MemberType NoteProperty -Name NetworkVIBVersion -Value $IOInfo.VibVersion 
            }
        if ($IOInfo.Count -gt 1){
            $vmhost | Add-Member -MemberType NoteProperty -Name NetworkDescription1 -Value $IOInfo.Description[0]
            $vmhost | Add-Member -MemberType NoteProperty -Name NetworkDrv1 -Value $IOInfo.Driver[0]
            $vmhost | Add-Member -MemberType NoteProperty -Name NetworkDrvVer1 -Value $IOInfo.DriverVersion[0]
            $vmhost | Add-Member -MemberType NoteProperty -Name NetworkFirmware1 -Value $IOInfo.FirmwareVersion[0]
            $vmhost | Add-Member -MemberType NoteProperty -Name NetworkVIBVersion1 -Value $IOInfo.VibVersion[0] 
            $vmhost | Add-Member -MemberType NoteProperty -Name NetworkDescription2 -Value $IOInfo.Description[2]
            $vmhost | Add-Member -MemberType NoteProperty -Name NetworkDrv2 -Value $IOInfo.Driver[2]
            $vmhost | Add-Member -MemberType NoteProperty -Name NetworkDrvVer2 -Value $IOInfo.DriverVersion[2]
            $vmhost | Add-Member -MemberType NoteProperty -Name NetworkFirmware2 -Value $IOInfo.FirmwareVersion[2]
            $vmhost | Add-Member -MemberType NoteProperty -Name NetworkVIBVersion2 -Value $IOInfo.VibVersion[2] 
            }
        $vmhost | Add-Member -MemberType NoteProperty -Name Product -Value $vmview.Config.Product.Name
        $vmhost | Add-Member -MemberType NoteProperty -Name Version -Value $vmview.Config.Product.Version
        $vmhost | Add-Member -MemberType NoteProperty -Name Build -Value $vmview.Config.Product.Build
        $vmhost | Add-Member -MemberType NoteProperty -Name ManagementIP -Value ($vmview.Config.Network.VNic.Spec | Where {$_.PortGroup -eq "Management Network"}).IP.IPAddress
        $vmhost | Add-Member -MemberType NoteProperty -Name Cluster -Value ($vmview.parent | Get-VIObjectByVIView | Select -ExpandProperty Name)
        #Clean out extra data from ( if it exists as it will not resolve parent with it
        $vmClusterName = $vmview.parent | Get-VIObjectByVIView | Select -ExpandProperty Name
        $vmClusterName = $vmClusterName.split('(')[0]
        #Get Host Cluster Parent folder If Host is returned assume no Folder and blank the return value
        $vmCluster = get-view -ViewType ClusterComputeResource -filter @{"Name" = ($vmClusterName)} 
        $hostFolder = $vmCluster.parent[0] | Get-VIObjectByVIView | Select -ExpandProperty Name
        if ($hostFolder -eq "host"){
            $vmhost | Add-Member -MemberType NoteProperty -Name Folder -Value ""
            } Else {
                $vmhost | Add-Member -MemberType NoteProperty -Name Folder -Value $hostFolder
        }
        #Null out Folder related variables for next loop throw
        $vmCluster = $null
        $hostFoler = $null
        #Get the DataCenter Name
        $vmhost | Add-Member -MemberType NoteProperty -Name DataCenter -Value (Get-Datacenter -VMHost $vmview.name | Select -ExpandProperty Name)
        ForEach ($CustomAttribute in $vmview.AvailableField){
            $vmhost | Add-Member -MemberType NoteProperty -Name $CustomAttribute.Name -Value ($vmview.Summary.CustomValue | ? {$_.Key -eq $CustomAttribute.Key}).value
        }
        #Add record to HostList       
        $Global:HostList += $vmhost
        $Count++
    }
}
#*************************
# EndFunction Get-HostList
#*************************


#**********************
# Function Write-to-CSV
#**********************
Function Write-to-CSV {
    [CmdletBinding()]
    Param()
    "Writing HostList from $Global:VCname to CSV"
    $Global:HostList | Export-CSV -Path $Global:Folder\$Global:VCname-HostList.csv -NoTypeInformation


}
#*************************
# EndFunction Write-to-CSV
#*************************


#**************************
# Function Convert-To-Excel
#**************************
Function Convert-To-Excel {
    [CmdletBinding()]
    Param()
   "Converting HostList from $Global:VCname to Excel"
    $workingdir = $Global:Folder+ "\*.csv"
    $csv = dir -path $workingdir

    foreach($inputCSV in $csv){
        $outputXLSX = $inputCSV.DirectoryName + "\" + $inputCSV.Basename + ".xlsx"
        ### Create a new Excel Workbook with one empty sheet
        $excel = New-Object -ComObject excel.application 
        $excel.DisplayAlerts = $False
        $workbook = $excel.Workbooks.Add(1)
        $worksheet = $workbook.worksheets.Item(1)

        ### Build the QueryTables.Add command
        ### QueryTables does the same as when clicking "Data » From Text" in Excel
        $TxtConnector = ("TEXT;" + $inputCSV)
        $Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
        $query = $worksheet.QueryTables.item($Connector.name)


        ### Set the delimiter (, or ;) according to your regional settings
        ### $Excel.Application.International(3) = ,
        ### $Excel.Application.International(5) = ;
        $query.TextFileOtherDelimiter = $Excel.Application.International(5)

        ### Set the format to delimited and text for every column
        ### A trick to create an array of 2s is used with the preceding comma
        $query.TextFileParseType  = 1
        $query.TextFileColumnDataTypes = ,2 * $worksheet.Cells.Columns.Count
        $query.AdjustColumnWidth = 1

        ### Execute & delete the import query
        $query.Refresh()
        $query.Delete()

        ### Get Size of Worksheet
        $objRange = $worksheet.UsedRange.Cells 
        $xRow = $objRange.SpecialCells(11).ow
        $xCol = $objRange.SpecialCells(11).column

        ### Format First Row
        $RangeToFormat = $worksheet.Range("1:1")
        $RangeToFormat.Style = 'Accent1'

        ### Save & close the Workbook as XLSX. Change the output extension for Excel 2003
        $Workbook.SaveAs($outputXLSX,51)
        $excel.Quit()
    }
    ## To exclude an item, use the '-exclude' parameter (wildcards if needed)
    remove-item -path $workingdir 

}
#*****************************
# EndFunction Convert-To-Excel
#*****************************

#***************
# Execute Script
#***************

# Get Start Time
$startDTM = (Get-Date)

#CLS
$ErrorActionPreference="SilentlyContinue"

"=========================================================="
" "
Write-Host "Get CIHS credentials" -ForegroundColor Yellow
$Global:Creds = Get-Credential -Credential $null

Get-VCenter
Connect-VC
"----------------------------------------------------------"
Verify-Folders
"----------------------------------------------------------"
Get-HostList
"----------------------------------------------------------"
Write-to-CSV
"----------------------------------------------------------"
Convert-To-Excel
"----------------------------------------------------------"
Disconnect-VC
"Open Explorer to $Global:Folder"
Invoke-Item $Global:Folder
Clean-Up

# Get End Time
$endDTM = (Get-Date)

# Echo Time elapsed
"Elapsed Time: $(($endDTM-$startDTM).totalseconds) seconds"

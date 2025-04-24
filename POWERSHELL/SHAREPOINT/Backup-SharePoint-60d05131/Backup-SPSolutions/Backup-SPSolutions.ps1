<#
.Synopsis
   SharePoint 2010 farm Solutions Back up
.DESCRIPTION
   This Script will help SP Admins to back up all WSP files in Folder as a backup.
   We can deploy directly from the back up or can be used as required.
.EXAMPLE
   Backup-SPSolutions -FolderPath C:\Temp -FolderName 'SPFarmSolutions'
.Contact
   chendrayan.exchange@hotmail.com
#>
function Backup-SPSolutions
{
    [CmdletBinding()]
    Param
    (
        # Provide Folder path to create a Back up folder 
        [Parameter(Mandatory=$true,
                   helpmessage="Enter the valid path",
                   Position=0)]
        [System.String]
        $FolderPath,

        # name the back up folder eg: SPFarmSolutionsBackup
        [System.String]
        $FolderName
    )

    Begin
    {
        Write-Host "Backing Up SharePoint Farm Solutions..." -ForegroundColor Yellow 
        New-Item $FolderPath\$FolderName -ItemType Directory -Force  
        Set-Location $FolderPath\$FolderName
        Start-Sleep 2
         
    }
    
    Process
    {
        (Get-SPFarm).Solutions | %{$Solutions = (Get-Location).Path + “\” + $_.Name; $_.SolutionFile.SaveAs($Solutions)}  
    }
    End
    {
        Write-Host "SharePoint Farm Solutions are backed up...." -ForegroundColor Yellow
        Invoke-Item $FolderPath\$FolderName
    }
}

Backup-SPSolutions -FolderPath C:\Temp -FolderName 'SPFarmSolutions'
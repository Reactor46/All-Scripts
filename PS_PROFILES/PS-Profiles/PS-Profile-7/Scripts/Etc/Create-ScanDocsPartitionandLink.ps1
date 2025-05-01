Install-Module NTFSSecurity -Confirm:$false -Force
Import-Module NTFSSecurity

# Find DVD-ROM & Change Drive Letter if D
$cdromdrive = (get-wmiobject win32_CDROMDrive).Drive
if ( $cdromdrive -match "D" )
    {
    $CDdrv = Get-WmiObject win32_volume -filter 'DriveLetter = "D:"'
    $CDdrv.DriveLetter = "E:"
    $CDdrv.Put() | out-null
    }

# Check to see if D Drive already exists
$Dbool = Get-Volume -DriveLetter D -ErrorAction SilentlyContinue
if ( $Dbool -eq $null )
    {
    # Resize C Drive and Create D Drive
    Function Get-FreeDiskSpace
        {
        $FullDiskSize = (Get-Partition -DriveLetter C | Get-Disk).size
        $FreeDiskSpace = $FullDiskSize
        Get-Partition | % { $FreeDiskSpace = $FreeDiskSpace - $_.Size }
        Return $FreeDiskSpace
        }

    $CPartitionSize = (Get-Partition -DriveLetter C).size
    $NewCPartionSize = (Get-FreeDiskSpace) + $CPartitionSize - 104857600
    Resize-Partition -DriveLetter C -Size $NewCPartionSize
    $DPartition = New-Partition -UseMaximumSize -DiskNumber (get-disk).Number 
    $Dpartition | Format-Volume -NewFileSystemLabel ScanDocs -Confirm:$false
    $DPartition | Set-Partition -NewDriveLetter D

    # Create Symlink for Scandocs
    mkdir C:\ScanDocs
    cmd /c "mklink /D D:\ScanDocs C:\ScanDocs"
    mkdir C:\Photos
    cmd /c "mklink /D D:\Photos C:\Photos"
    # Set Permissions on D Drive
    Get-Item D:\ | Get-NTFSAccess -Account 'NT AUTHORITY\Authenticated Users' | Remove-NTFSAccess
    Get-Item D:\ | Add-NTFSAccess -Account 'NT AUTHORITY\Authenticated Users' -AccessRights ReadAndExecute
    }
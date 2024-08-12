
$colRights = [System.Security.AccessControl.FileSystemRights]"Fullcontrol" 

$InheritanceFlag = [System.Security.AccessControl.InheritanceFlags]::None 
$PropagationFlag = [System.Security.AccessControl.PropagationFlags]::None 

$objType =[System.Security.AccessControl.AccessControlType]::Allow 

$objUser = New-Object System.Security.Principal.NTAccount("Winserv0\Administrateur") 

$objACE = New-Object System.Security.AccessControl.FileSystemAccessRule `
    ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 

$objACL = Get-ACL "C:\Backup\test.csv" 
$objACL.AddAccessRule($objACE) 

Set-ACL "C:\Backup\test.csv" $objACL
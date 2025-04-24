#
# This snippet gets a dir from ther user, sets owner of thet folder and subfolders to the
# Administrators group and grants the user IUSR Full access recursively.
#
#

#Stop IIS
IISReset /STOP

echo ""
echo "=======Set Folder Perms======="
echo ""
echo ""
echo "Set Administrators as Owner"
echo ""


foreach ($i in $args)
{
    icacls "$i" /setowner administrators /t
}

echo ""
echo "Grant IUSR R\W"
echo ""


foreach ($i in $args)
{
	icacls "$i" /grant:r IUSR:F /t
	icacls "$i" /grant:r IUSR:"(OI)(CI)"F /t
}

#Start IIS
IISReset /START


<# 
This script adds SA accounts for IIS 7
#>


$saccounts="IUSR","IIS_IUSRS"

$input=read-host "Enter base path"
foreach ($path in (Get-ChildItem $input))
    {
    Write-Host $path
    foreach ($account in $saccounts)
    {
    try
        {
            $colRights = [System.Security.AccessControl.FileSystemRights]"ReadAndExecute, Synchronize" 

            $InheritanceFlag = [System.Security.AccessControl.InheritanceFlags]::ContainerInherit `
                    -bor [System.Security.AccessControl.InheritanceFlags]::ObjectInherit
            $PropagationFlag = [System.Security.AccessControl.PropagationFlags]::None 

            $objType =[System.Security.AccessControl.AccessControlType]::Allow 

            $objUser = New-Object System.Security.Principal.NTAccount($account) 

            $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule `
                ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 

            $objACL = Get-ACL $path.fullname
            $objACL.AddAccessRule($objACE) 

            Set-ACL $path.fullname $objACL
        if ($account -eq "IUSR")
            {
                #fix this D.R.Y.
                $colRights = [System.Security.AccessControl.FileSystemRights]"Write" 
                $objType =[System.Security.AccessControl.AccessControlType]::Deny

                $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule `
                    ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 

                $objACL = Get-ACL $path.fullname
                $objACL.AddAccessRule($objACE) 

                Set-ACL $path.fullname$objACL
            }
        else
            {
                continue
            }

         }
    
        catch
        {
            continue
        }
    }
}
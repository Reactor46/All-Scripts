#########################################################################
#		Author: Vikas SUkhija
#               Description: Create Home folder
#               date: 09/17/2014
#               
#########################################################################

$date = get-date -format d
# replace \ by -
$time = get-date -format t
$month = get-date 
$month1 = $month.month
$year1 = $month.year

$date = $date.ToString().Replace(“/”, “-”)

$time = $time.ToString().Replace(":", "-")
$time = $time.ToString().Replace(" ", "")


$logs = ".\" + "Powershell" + $date + "_" + $time + "_.txt"

start-transcript $logs

# ListDirectory, ReadData, WriteData 
# CreateFiles, CreateDirectories, AppendData 
# ReadExtendedAttributes, WriteExtendedAttributes, Traverse
# ExecuteFile, DeleteSubdirectoriesAndFiles, ReadAttributes 
# WriteAttributes, Write, Delete 
# ReadPermissions, Read, ReadAndExecute 
# Modify, ChangePermissions, TakeOwnership
# Synchronize, FullControl

If ((Get-PSSnapin | where {$_.Name -match "Quest.ActiveRoles.ADManagement"}) -eq $null)
{
	Add-PSSnapin Quest.ActiveRoles.ADManagement
}

###################Define Variables####

$NetPath = "\\labnas\users$"
$users = get-content .\users.txt
$dletter = "H:"

#######################################

$users | foreach-object{

$qaduser = get-qaduser $_

$userhomepath = $NetPath + "\" + $_

	if(-not(Test-Path $userhomepath))
	{

	New-Item -Path $userhomepath -ItemType Directory
        Write-host "$userhomepath ------Created" -foregroundcolor Blue
        $acl = get-acl $userhomepath

	$inheritanceFlags = ([Security.AccessControl.InheritanceFlags]::ContainerInherit -bor `
                         [Security.AccessControl.InheritanceFlags]::ObjectInherit)
	$propagationFlags = [Security.AccessControl.PropagationFlags]::None


        $permissions = $_,"Modify",$inheritanceFlags,$propagationFlags,"Allow"
        $access = New-Object  system.security.accesscontrol.filesystemaccessrule($permissions)
	$acl.SetAccessRule($access)
	$acl | Set-Acl $userhomepath 
        $homedir = $qaduser.HomeDirectory
        	if ($homedir -like $null)
                   {
                    
                   Get-QADUser $_ | Set-QADUser -ObjectAttributes @{HomeDirectory = $userhomepath}
                   Get-QADUser $_ | Set-QADUser -ObjectAttributes @{HomeDrive = $dletter}
                   $usr = Get-QADUser $_

                   Write-host "Added Homedrive "$usr.HomeDrive" and Home directory "$usr.HomeDirectory""  -foregroundcolor Green
                   }
                 else {Write-host "$homedir already exists in AD for $_" -foregroundcolor yellow}
                 
        }
        else
        {

	Write-Warning -Message "'$userhomepath' already exists."
        $homedir = $qaduser.HomeDirectory
                if ($qaduser.HomeDirectory -like $null)
                   {
                    
                   Get-QADUser $_ | Set-QADUser -ObjectAttributes @{HomeDirectory = $userhomepath}
                   Get-QADUser $_ | Set-QADUser -ObjectAttributes @{HomeDrive = $dletter}
                   $usr = Get-QADUser $_

                   Write-host "Added Homedrive "$usr.HomeDrive" and Home directory "$usr.HomeDirectory""  -foregroundcolor Green
                   }
                 else {Write-host "$homedir already exists in AD for $_" -foregroundcolor yellow}
                 

       }
}


Stop-transcript
########################################################################
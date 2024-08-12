<#

    Author:         Panhead93
    Description:    A simple script that looks through all the folders and subfolders in a directory and lists all the folders that a specified user or group has access to.
                    Note that you don't have to type in the full name but it is recommended so that it can search using the most unique name.
                    The script will save the list of directories that the specified user or group has permissions to it finds to your Desktop.

#>


#Name of the group or user
$GroupOrUserName = Read-Host -Prompt "Name of the user or group"

#Path to the folder you would like to scan
$Directory = Read-Host -Prompt "Directory"

#Gets the path to your desktop
$DesktopPath = [Environment]::GetFolderPath("Desktop")

if (Test-Path $Directory)
{
    "Searching for all the folders within '$Directory' that $GroupOrUserName has access to" | Out-File "$DesktopPath\Folder Permissions.txt"
    
    Get-ChildItem $Directory -Recurse | Where-Object {$_.psiscontainer} | Where-Object {(Get-Acl $_.FullName).Access | Where-Object {$_.IdentityReference -like "*$GroupOrUserName*"} } | Format-Table FullName -AutoSize | Out-File "$DesktopPath\Folder Permissions.txt" -Append
    Write-Output('Saved the result in "Folder permissions.txt" on your desktop')
    Start-Sleep -Seconds 5
}
else
{
    Write-Output("No such directory! Exiting now...")
    Start-Sleep -Milliseconds 3500
}
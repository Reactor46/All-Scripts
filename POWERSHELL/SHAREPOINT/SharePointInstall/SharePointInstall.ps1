##
#This File Performs an Unattended Installation of SharePoint Server 2010
##

##
#Begin Setting Variables
##

$PathToConfigXML = '"C:\Install SharePoint\config.xml"'
$BinaryPath = "D:"

##
#Begin Script Execution
##

##
#Register Functions
##

#This function ensures that the path entered ends with a backslash.  We'll use this to build paths for executables we're installing

function EnsureBinaryPath
{
If ($BinaryPath.Endswith("\"))
    {
    }
else
    {
    Set-Variable -Name BinaryPath -Value "$BinaryPath\" -Scope 1
    }
}

#Call the function that ensures the BinaryPath variable ends with a backslash
EnsureBinaryPath

#Build the 2 commands we use to install the pre-requisites and to install SharePoint
$PreReqInstall = $BinaryPath + "PreRequisiteInstaller.exe /unattended"
$SharePointInstall = $BinaryPath + "Setup.exe /config $PathToConfigXML"

cls
#Call the installations and update the user of the installation progress
Write-Progress -Activity "Installing SharePoint Server 2010" -Status "Installing Pre-Requisites"
cmd.exe /c $PreReqInstall
Write-Progress -Activity "Installing SharePoint Server 2010" -Status "Installing SharePoint Server 2010" 
cmd.exe /C $SharePointInstall
Write-Progress -Activity "Installing SharePoint Server 2010" -Status "Completing SharePoint Server 2010 Installation"
Sleep 5
exit
#Get-ExchangeServerPlus.ps1
#v1.2, 09/17/2010
#Written By Paul Flaherty, blogs.flaphead.com
#Modified by Jeff Guillet, www.expta.com


#Get a list of Exchange servers in the Org excluding Edge servers
$MsxServers = Get-ExchangeServer | where {$_.ServerRole -ne "Edge"} | sort Name

#Loop through each Exchange server that is found
ForEach ($MsxServer in $MsxServers)
{

	#Get Exchange server version
	$MsxVersion = $MsxServer.ExchangeVersion

	#Create "header" string for output
	# Servername [Role] [Edition] Version Number
	$txt1 = $MsxServer.Name + " [" + $MsxServer.ServerRole + "] [" + $MsxServer.Edition + "] " + $MsxServer.AdminDisplayVersion #$MsxVersion.ExchangeBuild.toString()
	write-host $txt1

	#Connect to the Server's remote registry and enumerate all subkeys listed under "Patches"
	$Srv = $MsxServer.Name
	$key = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\461C2B4266EDEF444B864AD6D9E5B613\Patches\"
	$type = [Microsoft.Win32.RegistryHive]::LocalMachine
	$regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($type, $Srv)
	$regKey = $regKey.OpenSubKey($key)

	#Loop each of the subkeys (Patches) and gather the Installed date and Displayname of the Exchange 2007 patch
	$ErrorActionPreference = "SilentlyContinue"
	ForEach($sub in $regKey.GetSubKeyNames())
	{
		Write-Host "- " -nonewline
		$SUBkey = $key + $Sub
		$SUBregKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($type, $Srv)
		$SUBregKey = $SUBregKey.OpenSubKey($SUBkey)

		ForEach($SubX in $SUBRegkey.GetValueNames())
		{
			# Display Installed date and Displayname of the Exchange 2007 patch
			IF ($Subx -eq "Installed")   {
				$d = $SUBRegkey.GetValue($SubX)
				$d = $d.substring(4,2) + "/" + $d.substring(6,2) + "/" + $d.substring(0,4)
				write-Host $d -NoNewLine
			}
			IF ($Subx -eq "DisplayName") {write-Host ": "$SUBRegkey.GetValue($SubX)}
		}
	}
		write-host ""
}
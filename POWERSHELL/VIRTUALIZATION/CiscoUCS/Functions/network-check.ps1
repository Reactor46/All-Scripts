# (c) vbeninco@cisco.com
# 
# Get Packets (bytes/30seconds) going in and out of blades
# Note:  Please make sure your collection policy is set to collect every 30 seconds.
# Get the parameters 
param(	[string]$server = "10.93.234.241",
	[string]$user = "admin",
	[string]$passwd = "Cisco.123"
)

# make sure we have the Cisco module imported
# TODO check that its not already included.
Import-Module CiscoUCSPS

# take in the password.  
$pass = ConvertTo-SecureString $passwd -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential($user, $pass)

# only do one UCS right now... so many places we could improve this.
$handle = Connect-Ucs $server -Credential $cred

# get all the blades in the domain.
$blades = get-UcsBlade

# go through each blade.
foreach($blade in $blades) {
	Write-Host "Blade: "$blade.dn
	if ($blade.association -eq "Associated"){
		$sp = Get-UcsServiceProfile -Dn $blade.assignedToDn
		Write-Host "Service Profile: "$sp.dn
		$vnics = Get-UcsVnic -ServiceProfile $sp
		foreach($vnic in $vnics){
			$stats = Get-UcsAdaptorVnicStats -Dn ($vnic.equipmentdn + "/vnic-stats")
			Write-Host "`tVnic: "$vnic.rn" Tx: "$('{0:N3}' -f ($stats.BytesTxDeltaAvg * 0.000000254))"Mbps / Rx: "$('{0:N3}' -f ($stats.BytesRxDeltaAvg * 0.000000254))"Mbps"
		}
	}
}

Disconnect-UCS

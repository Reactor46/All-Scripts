function Get-UcsDvsConfig { 
param ([string] $Ucs)

	#Fetch known global policies
	[Cisco.Ucs.ManagedObject[]] $gp = Get-UcsVmVcenter  -Ucs $Ucs
	[Cisco.Ucs.ManagedObject[]] $gp += Get-UcsVmFolder   -Ucs $Ucs
	[Cisco.Ucs.ManagedObject[]] $gp += Get-UcsVmVcenterFolder -Ucs $Ucs
	[Cisco.Ucs.ManagedObject[]] $gp += Get-UcsVmVcenterDataCenter  -Ucs $Ucs
	[Cisco.Ucs.ManagedObject[]] $gp += Get-UcsVmVcenterDvs  -Ucs $Ucs
	[Cisco.Ucs.ManagedObject[]] $gp += Get-UcsVmPortProfile  -Ucs $Ucs
	[Cisco.Ucs.ManagedObject[]] $gp += Get-UcsVmVCenterKeyRing  -Ucs $Ucs
	[Cisco.Ucs.ManagedObject[]] $gp += Get-UcsVmVcenterExtensionKey -Ucs $Ucs

	# Now return it to the caller 
	return ,$gp
}

if ($args.Length -ne 1)
{
	Write-Error "Invoke with one argument."
}
else
{
	Get-UcsDvsConfig $args[0]
}

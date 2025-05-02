function Get-UcsGlobalPolicy { 
param ([string] $Ucs)

	#Fetch known global policies
	[Cisco.Ucs.ManagedObject[]]$gp = Get-UcsTimezone -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsNtpServer -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsDns -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsDnsServer -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsHttp -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsHttps -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsTelnet -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsWebSessionLimit -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsCimXml -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsSnmp -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsSnmpTrap -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsSnmpUser -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsFaultPolicy -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsSysDebugAutoCorefileExportTarget -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsSyslogConsole -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsSyslog -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsSyslogMonitor -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsSyslogClient -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsSyslogSource -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsSyslogFile -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsCallhomeSource -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsCallhomeSmtp -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsCallhome -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsCallhomeProfile -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsCallhomePolicy -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsCallhomePeriodicSystemInventory -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsCallhomeTestAlert -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsNativeAuth -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsAuthDomain -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsAuthDomainDefaultAuth -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsRadiusProvider -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsRole -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsLocale -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsaaaOrg -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsLocalUser -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsUserRole -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsRadiusProvider -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsProviderGroup -ucs $Ucs
	[Cisco.Ucs.ManagedObject[]]$gp += Get-UcsProviderReference -ucs $Ucs

	[Cisco.Ucs.ManagedObject[]]$gp = @($gp | Where-Object { $_ -ne $Null })
	# Now return it to the caller 
	return ,$gp
}

if ($args.Length -ne 1)
{
	Write-Error "Invoke with one argument."
}
else
{
	Get-UcsGlobalPolicy $args[0]
}



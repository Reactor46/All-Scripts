#Find Missing Updates on all Windows Servers in AD
$ResultsPath = "C:\Scripts\Reports\Servers"

Get-ADComputer -Server vdcmaddc01.corp.vegas.com -Filter {Operatingsystem -Like 'Windows Server® 2008*' -and Enabled -eq 'true'} -Properties * | Select -ExpandProperty Name | Out-FileUtf8NoBom $ResultsPath\CORP\2008\Corp-2K8-Servers.txt  -Append
Get-ADComputer -Server vdcmaddc01.corp.vegas.com -Filter {Operatingsystem -Like 'Windows Server 2008 R2*' -and Enabled -eq 'true'} -Properties * | Select -ExpandProperty Name | Out-FileUtf8NoBom $ResultsPath\CORP\2008R2\Corp-2K8R2-Servers.txt  -Append
Get-ADComputer -Server vdcmaddc01.corp.vegas.com -Filter {Operatingsystem -Like 'Windows Server 2012*' -and Enabled -eq 'true'} -Properties * | Select -ExpandProperty Name | Out-FileUtf8NoBom $ResultsPath\CORP\2012\Corp-2K12-Servers.txt  -Append
Get-ADComputer -Server vdcmaddc01.corp.vegas.com -Filter {Operatingsystem -Like 'Windows Server 2012 R2*' -and Enabled -eq 'true'} -Properties * | Select -ExpandProperty Name | Out-FileUtf8NoBom $ResultsPath\CORP\2012R2\Corp-2K12R2-Servers.txt  -Append
Get-ADComputer -Server vdcmaddc01.corp.vegas.com -Filter {Operatingsystem -Like 'Windows Server 2016*' -and Enabled -eq 'true'} -Properties * | Select -ExpandProperty Name | Out-FileUtf8NoBom $ResultsPath\CORP\2016\Corp-2K16-Servers.txt  -Append

Get-Content "$ResultsPath\CORP\2008\Corp-2k8-Servers.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CORP\2008\Corp-2k8-Servers-Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CORP\2008\Corp-2k8-Servers-Dead.txt" -Append }}

  Get-Content "$ResultsPath\CORP\2008R2\Corp-2k8R2-Servers.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CORP\2008R2\Corp-2k8R2-Servers-Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\2008R2\CORP\Corp-2k8R2-Servers-Dead.txt" -Append }}

  Get-Content "$ResultsPath\CORP\2012\Corp-2k12-Servers.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CORP\2012\Corp-2k12-Servers-Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CORP\2012\Corp-2k12-Servers-Dead.txt" -Append }}

  Get-Content "$ResultsPath\CORP\2012R2\Corp-2k12R2-Servers.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CORP\2012R2\Corp-2k12R2-Servers-Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CORP\2012R2\Corp-2k12R2-Servers-Dead.txt" -Append }}

  Get-Content "$ResultsPath\CORP\2016\Corp-2k16-Servers.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CORP\2016\Corp-2k16-Servers-Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\CORP\2016\Corp-2k16-Servers-Dead.txt" -Append }}


  			
			$mbsacli = "C:\Scripts\MBSA\mbsacli.exe"
            $Options = '/catalog "C:\Scripts\MBSA\wsusscn2.cab" /nvc /rd "C:\Scripts\Reports\Servers\CORP\Scans" /o "%D% - %C% (%T%).xml" '
			$UpdateRegex = '\| (.+) \| Missing \| (.+) \| (.+)? \|'
			$CheckResult | where { $_ -match $UpdateRegex } | foreach { [pscustomobject]@{ 'KBNumber' = $matches[1]; 'Severity' = $matches[3]; 'Title' = $matches[2] } }
	
  $CorpServers2k8 =  "$ResultsPath\CORP\2008\Corp-2k8-Servers-Alive.txt"
  $CorpServers2k8R2 = "$ResultsPath\CORP\2008R2\Corp-2k8R2-Servers-Alive.txt"
  $CorpServers2k12 = "$ResultsPath\CORP\2012\Corp-2k12-Servers-Alive.txt"
  $CorpServers2k12R2 = "$ResultsPath\CORP\2012R2\Corp-2k12R2-Servers-Alive.txt"
  $CorpServers2k16 = "$ResultsPath\CORP\2016\Corp-2k16-Servers-Alive.txt"


& $mbsacli /listfile $CorpServers2k8 /catalog "C:\Scripts\MBSA\wsusscn2.cab" /wi /nvc /nd /n Password+IIS+OS+SQL /rd "C:\Scripts\Reports\Servers\CORP\Scans"
& $mbsacli /listfile $CorpServers2k8R2 /catalog "C:\Scripts\MBSA\wsusscn2.cab" /wi /nvc /nd /n Password+IIS+OS+SQL /rd "C:\Scripts\Reports\Servers\CORP\Scans"
& type "C:\Scripts\Reports\Servers\CORP\Scans\*.mbsa" >> "C:\Scripts\Reports\Servers\CORP\Scans\final.xml"

Add-Content "C:\Scripts\Reports\Servers\CORP\Scans\final.xml" -n<fullscan>
 </fullscan>
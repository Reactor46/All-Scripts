add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010;
$database = $args.get(0);
if ( !$database )
 {
 Write-Host "Message: Can't find ""database"" argument. Check documentation."
 exit 1
 }
$Error.Clear();
$i = 1;
#Get-MailboxDatabaseCopyStatus -Identity $database | ForEach-Object `
 { `
 $stat = 3; `
 if ( $_.Status -eq "Mounted" ) { $stat = 0;} `
 if ( $_.Status -eq "Healthy" ) { $stat = 1;} `
 if ( $_.Status -eq "Dismounted" ) { $stat = 2;} `
 Write-Host Message.$i":" Database: $_.Name status $_.Status; `
 Write-Host Statistic.$i":" $stat; `
 if ($i -eq 10 ) { break; } `
 $i++; `
 }
if ($Error.Count -ne 0) 
 {
 Write-Host "Message.1: $($Error[0])";
 exit 1;
 }
for ( ; $i -lt 10; $i++) {
 Write-Host Message.$i":" No more databases.;
 Write-Host Statistic.$i":" 0;
 }
exit 0;
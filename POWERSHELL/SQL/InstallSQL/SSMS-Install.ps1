$src1= "\\pmmr.com\dfs1\Install\DBA\MSSQL\SSMS\18.5\SSMS-Setup-ENU.exe"
$src=$src1 -replace ' ', '` '
#$destination =Read-host " Enter Destination Location"
#$rc=Copy-Item -Path $src -Destination $destination -Recurse -Force -ErrorAction Stop

try
{
  
    if (Test-Path $src1)
    {
        #Invoke-Expression "$smssInstaller $installFlags";
	    write-host "SSMS install starting"
	    $SSMScmd=$src+" /install /quiet /norestart -Wait"
	    Invoke-Expression $SSMScmd
	    while( $running -eq $true)
	    {
	        $running= ( ( get-process | where processName -eq "setup").Length -gt 0);
	        start-sleep -s 10
	    }
	    $process=[System.IO.Path]::GetFileNameWithoutExtension($src)
	    $nid = (Get-Process $process).id
        Wait-Process -id $nid
        return 0;
    }
    else
    {
        write-host "SSMS installtion setup file does not exist"
        exit;
    }

}
catch
{
    return 1;
}

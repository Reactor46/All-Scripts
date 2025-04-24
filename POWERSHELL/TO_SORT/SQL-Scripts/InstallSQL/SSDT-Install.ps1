$src1= "\\pmmr.com\dfs1\Install\DBA\MSSQL\SSIS\SSDT-Setup-ENU.exe"
$src=$src1 -replace ' ', '` '
#$destination =Read-host " Enter Destination Location"
#$rc=Copy-Item -Path $src -Destination $destination -Recurse -Force -ErrorAction Stop

try
{
    if (Test-Path $src1)
    {
        #invoke-Expression $smssInstallerLoc $installFlags;
	    write-host "SSDT install starting"
	    $SSDTcmd=$src+" /install [:INSTALLIS :INSTALLVSSQL] /quiet /norestart log.txt -Wait"
	
	    Invoke-Expression $SSDTcmd
	    while( $running -eq $true)
	    {
	        $running= ( ( get-process | where processName -eq "setup").Length -gt 0);
	        start-sleep -s 10
	    }
	    $process=[System.IO.Path]::GetFileNameWithoutExtension($src)
	    $nid = (Get-Process $process).id
	    Wait-Process -id $nid
	    #Start-Process $src1+"\SSDT-Setup-ENU.exe" '/install [:INSTALLIS :INSTALLVSSQL] /quiet /norestart /log log.txt' -Wait;
	    #SSDT-Setup-ENU.exe /INSTALLALL[:vsinstances] /passive /log log.txt /layout c:\vs2017ssdt INSTALLIS [:INSTALLIS :INSTALLVSSQL]	
    }
    else
    {
        write-host " SSDT installtion setup file not exists"
        exit;
    }
}
catch
{
    write-host "SDT install fails"
}

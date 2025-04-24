. ./function.ps1
[reflection.assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo")
[reflection.assembly]::LoadWithPartialName("Microsoft.SqlServer.SqlWmiManagement")
function set-namedpiestatenable 
{
$cn=$env:COMPUTERNAME
$instance=get-SQlinstances 
$mc = New-Object ('Microsoft.SQLServer.Management.SMO.WMI.ManagedComputer')"$cn"
#Write-Host "Verifying Named Pipes for the SQL Service Instance on  $cn for $instance..." -ForegroundColor Green
$uri = "ManagedComputer[@Name='$cn']/ ServerInstance[@Name='$instance']/ServerProtocol[@Name='Np']"
    $np = $mc.GetSmoObject($uri)
    if(!($np.IsEnabled))
    {
     $np.IsEnabled = $true
    $np.Alter()
write-host " The Named pipes enabled on $cn for instance $instance is `t " $np.IsEnabled -ForegroundColor green
 }
}
    
function get-SQlinstances 
{
$Cn=$env:COMPUTERNAME
$sqlserviceins=Get-WmiObject -ComputerName $Cn win32_service -EA stop | where {$_.name -Like 'MSSQLSERVER' -or $_.name -like 'MSSQL$*' -and $_.state -eq 'running'}  

[System.Collections.ArrayList]$instancenames =@{}
foreach ($instance in $sqlserviceins)
{
    if($instance.name -eq "MSSQLServer")
        {
            
            $instancenames +=$instance.name;
        }

        else
        {
            $str=$($instance.name);
            $str=$str.split('$')[1];
            $instancenames +=$str;

        }
    }
    return $instancenames
}

function copy-SSISfiles ([string]$SSISdrive1)
{

if (!(Test-Path ($SSISdrive1 + ":\SSIS")))
    {	
        #Create SSIS folder
        $SSISDir=$SSISdrive1+":" 
        new-item -path $SSISDir -name "\SSIS" -type directory -force
    }

    $loc=$SSISdrive1+":\"
    Copy-Item -Path 'C:\temp\InstallSQL\SSIS\' -Recurse -force -Destination $loc

}

function ctrlmact
{
    switch ( $env )
    {
        "Prod"   { $account = '[PMMR\wm_ControlM]' }
        "Test"   { $account = '[PMMR\wm_ControlM_Test]' }
        "Test02" { $account = '[PMMR\wm_ControlM_Test02]' }
        "Dev"    { $account = '[PMMR\wm_ControlM_Dev]' }
        "Dev02"  { $account = '[PMMR\wm_ControlM_Dev02]' }
        default  {"Environment variable not existed"; return}
    } 
    return $account
}


function executeScripts([string]$pPath, [string]$ISpath, $jobopdir, $sqlcmdpath)
{	
	copy-SSISfiles $ISpath
	set-namedpiestatenable 
	[Boolean]$successful = $true
	$isLocal=1
	$serverName=[Environment]::GetEnvironmentVariable("SQL_SERVERNAME","Machine")
	$env=[Environment]::GetEnvironmentVariable("SERVER_ENVIRONMENT","Machine")
	if(Test-Path $pPath)
	{
        $scripts = Get-ChildItem $pPath -Filter "*.sql" | Sort-Object -Property Name | Select FullName, Name
		if(!$scripts)
        {
            return
        }
		else
		{
			if (!(Test-Path ($jobopdir + ":\JobLog")))
			{	
				#Create JobLog folder
				$jobopdir1=$jobopdir+":" 
				new-item -path $jobopdir1 -name "\JobLog" -type directory -force
				#new-item -type directory -path $jobopdir1 -name ":\JobLog" -type directory -force
			}
			# if the ExecutedScripts folder does not exist, then create it (we have scripts to run)
	
			if (!(Test-Path ($pPath + "\ExecutedScripts")))
			{	
				#Create ExecutedScripts folder 
				new-item -path $pPath -name "\ExecutedScripts" -type directory -force
			}
		}
        
		#$serverName = $script:sqlName
		#write-host $serverName
		$ctrlmacct=ctrlmact
		write-host "ControlM account name is : "$ctrlmacct
		foreach($s in $scripts)
		{
			$script = $s.FullName 
			$sqlcmdfl="""$sqlcmdpath"""
			$returnVal = & cmd /c $sqlcmdfl -i $script -E -S $serverName -b -v ctlm="'$ctrlmacct'" ISDR="'$ISpath'" env="'$env'" joblogop="'$jobopdir'"
			#backuppath="'$bkppath'" env="'$env'" ISDR="'$ISpath'" joboutpufiledir="'$jobopdir'"
    		if($LASTEXITCODE -ne 0)
    		{
    			write-host "$script failed to execute"
			"Script $script faield with error as: $returnVal" | Out-File "C:\temp\installSQl\script_error.log" -Append	
    		}
    		else
    		{
			    write-host "$script executed" 	
               	$DestinationFileLocation = $s.FullName -replace $s.Name, ("ExecutedScripts\" + $s.Name)
				Move-Item -Path $s.FullName -Destination $DestinationFileLocation -Force;
		"**************************************************************************** `n" | Out-File "C:\temp\installSQl\scripts_output.log" -Append
		"Results of $script $returnVal" | Out-File "C:\temp\installSQl\scriptsoutput.log" -Append
		"****************************************************************************" | Out-File "C:\temp\installSQl\scripts_output.log" -Append
          
            }
        }
	}
	else
	{
	
		write-host $pPath $fail
		throw
	}
	
}

#$infrastructurePath= read-host " please enter the infrastructure path to execute .sql  like C:\temp\"
$infrastructurePath="C:\temp\InstallSQL\Scripts\"
$sqlcmdloc=Get-ChildItem 'C:\Program Files\' -recurse | where {$_.name -eq "sqlcmd.exe"} | sort LastWriteTime | select -Last 1
$sqlcmdlo=$sqlcmdloc.fullname
#$SSISFolderPath= Read-host " Enter SSIS drive letter to change in the SSIS jobs like E or F "
$SSISdrive= SSISDriveLetter
$joboutputloc= joboutputdriveletter
if(!(Test-Path ($SSISdrive+':')))
{
    write-host " SSIS  drive $SSISdrive : does not exist"
	exit;
};
if(!(Test-Path ($joboutputloc+':'))){
    write-host " Job Ouptut $joboutputloc : drive does not exist"
	exit;
};
write-host "Updating the permissiosns on Backup shared folder"
Update-SQLBackupFolderPermissions

executeScripts $infrastructurePath $SSISdrive $joboutputloc $sqlcmdlo

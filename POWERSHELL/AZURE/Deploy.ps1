<#
 * Copyright Microsoft Corporation
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
#>

$scriptFolder = Split-Path -Parent $MyInvocation.MyCommand.Definition
. "$scriptFolder\sharedfunctions.ps1"

################## Functions ##############################

function Deploy()
{
    $VerbosePreference = "SilentlyContinue"

    cls

    if((IsAdmin) -eq $false)
    {
        Write-Host "Must run PowerShell elevated."
        return
    }

    Import-Module `
        "C:\Program Files (x86)\Microsoft SDKs\Windows Azure\PowerShell\Azure\Azure.psd1" `
        -ErrorAction Stop
    
    Write-Host "Enabling PowerShell remoting and the CredSSP client on the local machine..."
    Enable-PSRemoting -Force -ErrorAction Stop | Out-Null
    Enable-WSManCredSSP -Role client -DelegateComputer "*.cloudapp.net" -Force -ErrorAction Stop | Out-Null
    $regKey = "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\Credssp\PolicyDefaults\AllowFreshCredentialsWhenNTLMOnlyDomain"
    Set-ItemProperty $regKey -Name WSMan -Value "WSMAN/*.cloudapp.net" -ErrorAction Stop | Out-Null
    $regKey = "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\Credssp\PolicyDefaults\AllowFreshCredentialsWhenNTLMOnly"
    Set-ItemProperty $regKey -Name WSMan -Value "WSMAN/*.cloudapp.net" -ErrorAction Stop | Out-Null

    $d = get-date
    Write-Host "Starting Deployment $d"

    #region initialize Azure variables
    $subscription = Get-AzureSubscription -Current -ErrorAction Stop

    if($subscription -eq $null)
    {
        Write-Host "Windows Azure Subscription is not configured or the specified subscription name is invalid."
        Write-Host "Use Get-AzurePublishSettingsFile and Import-AzurePublishSettingsFile first"
        return
    }

    while($true)
    {
        $storageAccountName = "contososa" + (RandomString)
        if((Test-AzureName -Storage $storageAccountName) -eq $true)
        {
            Write-Host "Dynamically generated $storageAccountName is in use. Looking for another."
        }
        else
        {
            Write-Host "Using $storageAccountName for storage account name"
            break
        }
    }

    $dcServiceName = GenerateServiceName
    $sqlServiceName = GenerateServiceName
    $clientServiceName = GenerateServiceName

    $winImageName = (Get-AzureVMImage | where {$_.ImageFamily -eq "Windows Server 2012 Datacenter"} | sort PublishedDate -Descending)[0].ImageName
    $sqlImageName = (Get-AzureVMImage | where {$_.ImageFamily -eq "SQL Server 2012 SP1 Enterprise on Windows Server 2012"} | sort PublishedDate -Descending)[0].ImageName

    [xml] $config = gc "$scriptFolder\Config\DeploymentConfig.xml"

    #endregion

    #region create Azure environment
    New-AzureAffinityGroup `
        -Name $config.Azure.AffinityGroup `
        -Location $config.Azure.Location `
        -ErrorAction Stop

    Set-AzureVNetConfig `
        -ConfigurationPath "$scriptFolder\Config\NetworkConfig.xml" `
        -ErrorAction Stop

    New-AzureStorageAccount `
        -StorageAccountName $storageAccountName `
        -AffinityGroup $config.Azure.AffinityGroup `
        -ErrorAction Stop |
            Out-Null

    Set-AzureSubscription `
        -SubscriptionName $subscription.SubscriptionName `
        -CurrentStorageAccount $storageAccountName `
        -ErrorAction Stop |
            Out-Null

    #endregion

    #region create and configure DC server
    $dc = $config.Azure.AzureVMGroups.VMRole | where {$_.Name -eq "DomainControllers"} 
    $password = GetPasswordByUserName $dc.ServiceAccountName $config.Azure.ServiceAccounts.ServiceAccount 
    Write-Host "Creating DC server $($dc.Name)"
    CreateVM `
        -vmName $dc.AzureVM.Name `
        -imageName $winImageName `
        -size $dc.AzureVM.VMSize `
        -subnetNames $dc.SubnetNames `
        -adminUserName $dc.ServiceAccountName `
        -password $password `
        -serviceName $dcServiceName `
        -newService $true `
        -vnetName $config.Azure.VNetName `
        -affinityGroup $config.Azure.AffinityGroup `
        -windowsDomain $false `
        -dataDisks $null

    $ad = $config.Azure.ActiveDirectory
    Write-Host "Creating domain $($ad.DnsDomain)"
    RunWSManScriptBlock `
        -serviceName $dcServiceName `
        -vmName $dc.AzureVM.Name `
        -userName $dc.ServiceAccountName `
        -password $password `
        -argumentList $ad.DnsDomain, $ad.Domain, $password `
        -scriptBlock `
        {
            param($FQDN, $domainName, $vmAdminPassword)
            dcpromo.exe `
                /unattend `
                /ReplicaOrNewDomain:Domain `
                /NewDomain:Forest `
                /NewDomainDNSName:$FQDN `
                /ForestLevel:4 `
                /DomainNetbiosName:$domainName `
                /DomainLevel:4 `
                /InstallDNS:Yes `
                /ConfirmGc:Yes `
                /CreateDNSDelegation:No `
                /DatabasePath:"C:\Windows\NTDS" `
                /LogPath:"C:\Windows\NTDS" `
                /SYSVOLPath:"C:\Windows\SYSVOL" `
                /SafeModeAdminPassword:"$vmAdminPassword" `
        }

    Write-Host " Wait for $($dc.AzureVM.Name) to reboot"
    Start-Sleep -Seconds 120

    Write-Host "Configuring domain accounts"
    $serialXML = New-Object System.Xml.XmlDocument  
    $serialXML.AppendChild($serialXML.ImportNode($config.Azure.ServiceAccounts, $true)) | Out-Null
    RunWSManScriptBlock `
        -serviceName $dcServiceName `
        -vmName $dc.AzureVM.Name `
        -userName $dc.ServiceAccountName `
        -password $password `
        -argumentList $serialXML `
        -scriptBlock `
        {
            param($accounts)

            Write-Host "Configure AD objects on $env:COMPUTERNAME..."

            Import-Module ActiveDirectory

            foreach($account in $accounts.ServiceAccounts.ServiceAccount)
            {
                if ($account.UserName.Contains('\') -and ([string]::IsNullOrEmpty($account.Create) -or (-not $account.Create.Equals('No'))))
                {
                    $uname = $account.UserName.Split('\')[1]
                    $password = ConvertTo-SecureString $account.Password -AsPlainText -Force
                    Write-Host " Add AD account" $account.UserName
                    New-ADUser `
                        -Name $uname `
                        -AccountPassword $password `
                        -PasswordNeverExpires $true `
                        -ChangePasswordAtLogon $false `
                        -Enabled $true

                    if($account.Type -eq "DBA")
                    {
                        Write-Host " Configure $($account.UserName) with needed WSFC permissions" 
                        Cd ad:
                        $sid = new-object System.Security.Principal.SecurityIdentifier (Get-ADUser $uname).SID
                        $guid = new-object Guid bf967a86-0de6-11d0-a285-00aa003049e2
                        $ace1 = new-object System.DirectoryServices.ActiveDirectoryAccessRule $sid,"CreateChild","Allow",$guid,"All"
                        $dn = (Get-ADDomain | select -ExpandProperty DistinguishedName )
                        $corp = Get-ADObject -Identity $dn
                        $acl = Get-Acl $corp
                        $acl.AddAccessRule($ace1)
                        Set-Acl -Path $dn -AclObject $acl
                    }
                }
            }
        }

    #endregion

    #region create and configure VMs for WSFC nodes
            
    $sql = $config.Azure.AzureVMGroups.VMRole | where {$_.Name -eq "SQLServers"} 
    $password = GetPasswordByUserName $sql.ServiceAccountName $config.Azure.ServiceAccounts.ServiceAccount 
    $DBAAcct = $config.Azure.ServiceAccounts.ServiceAccount | where {$_.Type -eq "DBA"}
    $domainAcct = $config.Azure.ServiceAccounts.ServiceAccount | where {$_.Type -eq "WindowsDomain"}
    $dnsIP = (Get-AzureVM -ServiceName $dcServiceName -Name $dc.AzureVM.Name).IpAddress
    foreach($vm in $sql.AzureVM)
    {
        if ($vm.Type -eq "Quorum")
        {
            $imageName = $winImageName
        }
        else 
        {
            $imageName = $sqlImageName
        }
        $newService = $false
        if ($sql.AzureVM.IndexOf($vm) -eq 0)
        {
            $newService = $true;
        }
    
        Write-Host "Creating $($vm.Name)..."
        # create the VM
        CreateVM `
            -vmName $vm.Name `
            -imageName $imageName `
            -size $vm.VMSize `
            -subnetNames $sql.SubnetNames `
            -adminUserName $sql.ServiceAccountName `
            -password $password `
            -serviceName $sqlServiceName `
            -newService $newService `
            -availabilitySet $sql.AvailabilitySet `
            -vnetName $config.Azure.VNetName `
            -affinityGroup $config.Azure.AffinityGroup `
            -windowsDomain $true `
            -dnsIP $dnsIP `
            -domainJoin $ad.DnsDomain `
            -domain $ad.Domain `
            -domainUserName $domainAcct.UserName `
            -domainPassword $domainAcct.Password

        Write-Host "initializing $($vm.Name)..."
        # initialize the VM
        RunWSManScriptBlock `
            -serviceName $sqlServiceName `
            -vmName $vm.Name `
            -userName $sql.ServiceAccountName `
            -password $password `
            -argumentList $DBAAcct.UserName `
            -scriptBlock `
            {
                param($domainAcct)

                Write-Host " Adding $domainAcct as local administrator"
                net localgroup administrators "$domainAcct" /Add

                If ((Get-Service -Name "MSSQLSERVER" -ErrorAction SilentlyContinue) -ne $null)
                {
                    Write-Host " *A SQL Server node. Initialize SQL Server..."
            
                    Set-ExecutionPolicy -Execution RemoteSigned -Force
                    Import-Module -Name "sqlps" -DisableNameChecking

                    Write-Host " Installing failover clustering features"

                    Import-Module ServerManager
                    Add-WindowsFeature `
                        'Failover-Clustering', `
                        'RSAT-Clustering-Mgmt', `
                        'RSAT-Clustering-PowerShell', `
                        'RSAT-Clustering-CmdInterface' |
                            Out-Null

                    Write-Host " Adding $domainAcct as sysadmin"
                    Invoke-SqlCmd -Query "EXEC sp_addsrvrolemember '$domainAcct', 'sysadmin'" -ServerInstance "."

                    Write-Host " Adding NT AUTHORITY\SYSTEM with required permissions"
                    Invoke-SqlCmd -Query "CREATE LOGIN [NT AUTHORITY\SYSTEM] FROM WINDOWS" -ServerInstance "."
                    Invoke-SqlCmd -Query "GRANT ALTER ANY AVAILABILITY GROUP TO [NT AUTHORITY\SYSTEM] AS SA" -ServerInstance "." 
                    Invoke-SqlCmd -Query "GRANT CONNECT SQL TO [NT AUTHORITY\SYSTEM] AS SA" -ServerInstance "."
                    Invoke-SqlCmd -Query "GRANT VIEW SERVER STATE TO [NT AUTHORITY\SYSTEM] AS SA" -ServerInstance "."

                    Write-Host " Opening firewall port 1433"
                    netsh advfirewall firewall add rule `
                        name='SQL Server (TCP-In)' `
                        program='C:\Program Files\Microsoft SQL Server\MSSQL11.MSSQLSERVER\MSSQL\Binn\sqlservr.exe' `
                        dir=in `
                        action=allow `
                        protocol=TCP |
                            Out-Null
        
                    Write-Host " Enable delegation of client credentials for PS remoting"
                    Enable-WSManCredSSP –Role Server -Force | Out-Null
                }

                Write-Host "Done with $env:COMPUTERNAME"
            }
    }
    #endregion

    #region configure AG
    Write-Host "Starting AG configuration"

    $server1 = ($sql.AzureVM | where {$_.Type -eq "Primary"}).Name
    $server2 = ($sql.AzureVM | where {$_.Type -eq "Secondary"}).Name
    $serverQuorum = ($sql.AzureVM | where {$_.Type -eq "Quorum"}).Name
    $cluster = $config.Azure.SQLCluster
    $acct1 = $cluster.PrimaryServiceAccountName
    $acct2 = $cluster.SecondaryServiceAccountName
    $password1 = GetPasswordByUserName $acct1 $config.Azure.ServiceAccounts.ServiceAccount
    $password2 = GetPasswordByUserName $acct2 $config.Azure.ServiceAccounts.ServiceAccount
    RunWSManScriptBlock `
        -serviceName $sqlServiceName `
        -vmName $server1 `
        -userName $DBAAcct.UserName `
        -password $DBAAcct.Password `
        -credSSP $true `
        -argumentList $server1, $acct1, $password1, $server2, $acct2, $password2 `
        -scriptBlock `
        {
            param($server1, $acct1, $password1, $server2, $acct2, $password2)

            $timeout = New-Object System.TimeSpan -ArgumentList 0, 0, 30

            # Import SQL Server PowerShell Provider
            Set-ExecutionPolicy RemoteSigned -Force
            Import-Module "sqlps" -DisableNameChecking

            Write-Host " Change the SQL Server service account for $server1 to $acct1"
            $wmi1 = new-object ("Microsoft.SqlServer.Management.Smo.Wmi.ManagedComputer") $server1
            $wmi1.services | where {$_.Type -eq 'SqlServer'} | foreach{$_.SetServiceAccount($acct1,$password1)}
            $svc1 = Get-Service -ComputerName $server1 -Name 'MSSQLSERVER'
            $svc1.Stop()
            $svc1.WaitForStatus([System.ServiceProcess.ServiceControllerStatus]::Stopped,$timeout)
            $svc1.Start(); 
            $svc1.WaitForStatus([System.ServiceProcess.ServiceControllerStatus]::Running,$timeout)

            Write-Host " Change the SQL Server service account for $server2 to $acct2"
            $wmi2 = new-object ("Microsoft.SqlServer.Management.Smo.Wmi.ManagedComputer") $server2
            $wmi2.services | where {$_.Type -eq 'SqlServer'} | foreach{$_.SetServiceAccount($acct2,$password2)}
            $svc2 = Get-Service -ComputerName $server2 -Name 'MSSQLSERVER'
            $svc2.Stop()
            $svc2.WaitForStatus([System.ServiceProcess.ServiceControllerStatus]::Stopped,$timeout)
            $svc2.Start(); 
            $svc2.WaitForStatus([System.ServiceProcess.ServiceControllerStatus]::Running,$timeout)
        }

    Write-Host "Creating the WSFC cluster"
    $clusterScript = "$scriptFolder\CreateAzureFailoverCluster.ps1"
    Unblock-File -Path $clusterScript -ErrorAction Stop
    RunWSManScriptFile `
        -serviceName $sqlServiceName `
        -vmName $server1 `
        -userName $DBAAcct.UserName `
        -password $DBAAcct.Password `
        -credSSP $true `
        -argumentList $cluster.Name, @($server1,$server2) `
        -scriptFile $clusterScript |
            Out-Null

    RunWSManScriptBlock `
        -serviceName $sqlServiceName `
        -vmName $serverQuorum `
        -userName $DBAAcct.UserName `
        -password $DBAAcct.Password `
        -argumentList $acct1, $acct2, ($ad.Domain + "\" + $cluster.Name) `
        -scriptBlock `
        {
            param($acct1, $acct2, $clusterAcct)

            $backupDir = "C:\backup"
            $quorumDir = "C:\quorum"

            #Create share folder for quorum configuration
            New-Item $quorumDir -ItemType directory | Out-Null
            net share quorum=$quorumDir "/grant:$clusterAcct$,FULL" | Out-Null
            icacls.exe "$quorumDir" /grant:r ("$clusterAcct$" + ":(OI)(CI)F") | Out-Null

            #Create backup directory and grant permissions for the SQL Server service accounts
            New-Item $backupDir -ItemType directory | Out-Null
            net share backup=$backupDir "/grant:$acct1,FULL" "/grant:$acct2,FULL" | Out-Null
            icacls.exe "$backupDir" /grant:r ("$acct1" + ":(OI)(CI)F") ("$acct2" + ":(OI)(CI)F") ("$acct3" + ":(OI)(CI)F") | Out-Null
        }

    RunWSManScriptBlock `
        -serviceName $sqlServiceName `
        -vmName $server1 `
        -userName $DBAAcct.UserName `
        -password $DBAAcct.Password `
        -credSSP $true `
        -argumentList $server1, $server2, $serverQuorum, $acct1, $acct2, $cluster.Name, $cluster.Database, $cluster.AvailabilityGroup `
        -scriptBlock `
        {
            param($server1, $server2, $serverQuorum, $acct1, $acct2, $clusterName, $DatabaseName, $AvailabilityGroupName)

            $timeout = New-Object System.TimeSpan -ArgumentList 0, 0, 30
            $backupShare = "\\$serverQuorum\backup"
            $quorumShare = "\\$serverQuorum\quorum"

            Write-Host " Set quorum to file share majority with $serverQuorum"
            Import-Module FailoverClusters
            Set-ClusterQuorum –NodeAndFileShareMajority $quorumShare | Out-Null

            # Import SQL Server PowerShell Provider
            Set-ExecutionPolicy RemoteSigned -Force
            Import-Module "sqlps" -DisableNameChecking

            Write-Host " Enable AlwaysOn Availability Groups for $server1 and $server2"
            Enable-SqlAlwaysOn `
                -Path SQLSERVER:\SQL\$server1\Default `
                -Force
            Enable-SqlAlwaysOn `
                -Path SQLSERVER:\SQL\$server2\Default `
                -NoServiceRestart
            $svc2 = Get-Service -ComputerName $server2 -Name 'MSSQLSERVER'
            $svc2.Stop()
            $svc2.WaitForStatus([System.ServiceProcess.ServiceControllerStatus]::Stopped,$timeout)
            $svc2.Start(); 
            $svc2.WaitForStatus([System.ServiceProcess.ServiceControllerStatus]::Running,$timeout) 

            Write-Host " Create a directory for backups and restores"
            $backup = "C:\backup"
            New-Item $backup -ItemType directory | Out-Null
            net share backup=$backup "/grant:$acct1,FULL" "/grant:$acct2,FULL" | Out-Null
            icacls.exe "$backup" /grant:r ("$acct1" + ":(OI)(CI)F") ("$acct2" + ":(OI)(CI)F") | Out-Null

            Write-Host " Create database $DatabaseName, and restore its backups on $server2 with NO RECOVERY"
            Invoke-SqlCmd -Query "CREATE database $DatabaseName"
            Backup-SqlDatabase -Database $DatabaseName -BackupFile "$backupShare\db.bak" -ServerInstance $server1
            Backup-SqlDatabase -Database $DatabaseName -BackupFile "$backupShare\db.log" -ServerInstance $server1 -BackupAction Log
            Restore-SqlDatabase -Database $DatabaseName -BackupFile "$backupShare\db.bak" -ServerInstance $server2 -NoRecovery
            Restore-SqlDatabase -Database $DatabaseName -BackupFile "$backupShare\db.log" -ServerInstance $server2 -RestoreAction Log -NoRecovery 

            Write-Host " Create the availability group"
            $endpoint = 
                New-SqlHadrEndpoint MyMirroringEndpoint `
                -Port 5022 `
                -Path "SQLSERVER:\SQL\$server1\Default"
            Set-SqlHadrEndpoint `
                -InputObject $endpoint `
                -State "Started" |
                    Out-Null
            $endpoint = 
                New-SqlHadrEndpoint MyMirroringEndpoint `
                -Port 5022 `
                -Path "SQLSERVER:\SQL\$server2\Default"
            Set-SqlHadrEndpoint `
                -InputObject $endpoint `
                -State "Started" |
                    Out-Null

            Invoke-SqlCmd -Query "CREATE LOGIN [$acct2] FROM WINDOWS" -ServerInstance $server1
            Invoke-SqlCmd -Query "GRANT CONNECT ON ENDPOINT::[MyMirroringEndpoint] TO [$acct2]" -ServerInstance $server1
            Invoke-SqlCmd -Query "CREATE LOGIN [$acct1] FROM WINDOWS" -ServerInstance $server2
            Invoke-SqlCmd -Query "GRANT CONNECT ON ENDPOINT::[MyMirroringEndpoint] TO [$acct1]" -ServerInstance $server2 

            $primaryReplica = 
                New-SqlAvailabilityReplica `
                -Name $server1 `
                -EndpointURL "TCP://$server1.corp.contoso.com:5022" `
                -AvailabilityMode "SynchronousCommit" `
                -FailoverMode "Automatic" `
                -Version 11 `
                -AsTemplate
            $secondaryReplica = 
                New-SqlAvailabilityReplica `
                -Name $server2 `
                -EndpointURL "TCP://$server2.corp.contoso.com:5022" `
                -AvailabilityMode "SynchronousCommit" `
                -FailoverMode "Automatic" `
                -Version 11 `
                -AsTemplate

            New-SqlAvailabilityGroup `
                -Name $AvailabilityGroupName `
                -Path "SQLSERVER:\SQL\$server1\Default" `
                -AvailabilityReplica @($primaryReplica,$secondaryReplica) `
                -Database $DatabaseName |
                    Out-Null
            Join-SqlAvailabilityGroup `
                -Path "SQLSERVER:\SQL\$server2\Default" `
                -Name $AvailabilityGroupName
            Add-SqlAvailabilityDatabase `
                -Path "SQLSERVER:\SQL\$server2\Default\AvailabilityGroups\$AvailabilityGroupName" `
                -Database $DatabaseName 
        }

    Write-Host "Create the availability group listener"
    # create AG listener
    $listenerScript = "$scriptFolder\ConfigureAGListenerCloudOnly.ps1"
    Unblock-File -Path $listenerScript -ErrorAction Stop
    & $listenerScript `
        -AGName $cluster.AvailabilityGroup `
        -ListenerName $cluster.Listener `
        -ServiceName $sqlServiceName `
        -EndpointName "SQLEndpoint" `
        -EndpointPort "1433" `
        -WSFCNodes $server1, $server2 `
        -DomainAccount $DBAAcct.UserName `
        -Password $DBAAcct.Password

    Write-Host "Done with AG configuration on ContosoSQL1"

    #endregion

    $d = get-date
    Write-Host "End Deployment $d"

    #region test listener connectivity with client VM

    $client = $config.Azure.AzureVMGroups.VMRole | where {$_.Name -eq "SQLClient"} 
    $password = GetPasswordByUserName $client.ServiceAccountName $config.Azure.ServiceAccounts.ServiceAccount 

    Write-Host "Creating a VM in a separate cloud service to test listener (not reachable within the same cloud service)..."
    CreateVM `
        -vmName $client.Name `
        -imageName $sqlImageName `
        -size $client.AzureVM.VMSize `
        -subnetNames $client.SubnetNames `
        -adminUserName $client.ServiceAccountName `
        -password $password `
        -serviceName $clientServiceName `
        -newService $true `
        -vnetName $config.Azure.VNetName `
        -affinityGroup $config.Azure.AffinityGroup `
        -windowsDomain $true `
        -dnsIP $dnsIP `
        -domainJoin $ad.DnsDomain `
        -domain $ad.Domain `
        -domainUserName $domainAcct.UserName `
        -domainPassword $domainAcct.Password `
        -dataDisks $null

    RunWSManScriptBlock `
        -serviceName $clientServiceName `
        -vmName $client.Name `
        -userName $client.ServiceAccountName `
        -password $password `
        -argumentList $DBAAcct.UserName `
        -scriptBlock `
        {
            param($DBA)
            
            net localgroup administrators "$DBA" /Add | Out-Null
            Enable-WSManCredSSP –Role Server -Force | Out-Null
        }

    Write-Host "Testing connectivity to availability group listener"
    RunWSManScriptBlock `
        -serviceName $clientServiceName `
        -vmName $client.Name `
        -userName $DBAAcct.UserName `
        -password $DBAAcct.Password `
        -credSSP $true `
        -argumentList $cluster.Listener, $cluster.Database `
        -scriptBlock `
        {
            param($ListenerName, $DatabaseName)
            
            Write-Host "(Successful query indicates connectivity to listener)"
            sqlcmd -S $ListenerName -d $DatabaseName -Q "select @@servername, db_name()" -l 15
        }

    #endregion

} # end Main()

function RandomString ($length = 6)
{

    $digits = 48..57
    $letters = 65..90 + 97..122
    $rstring = get-random -count $length `
            -input ($digits + $letters) |
                    % -begin { $aa = $null } `
                    -process {$aa += [char]$_} `
                    -end {$aa}
    return $rstring.ToString().ToLower()
}

function GenerateServiceName()
{
    while($true)
    {
        $serviceName = "ContosoAG-" + (RandomString)
        if((Test-AzureName -Service $serviceName) -ne $true)
        {
            Write-Host "Using $serviceName for cloud service name"
            break
        }
    }

    return $serviceName
}

function CreateVM (
$vmName, 
$imageName, 
$size,
$subnetNames,
$adminUserName,
$password,
$serviceName,
$newService = $false,
$availabilitySet = "",
$vnetName,
$affinityGroup,
$windowsDomain = $false,
$dnsIP,
$domainJoin,
$domain,
$domainUserName,
$domainPassword,
$dataDisks = @()
)
{
    if($availabilitySet -ne "")
    {
    	$vmConfig = New-AzureVMConfig -Name $vmName -InstanceSize $size -ImageName $imageName -AvailabilitySetName $availabilitySet -ErrorAction Stop
    }
    else
    {
    	$vmConfig = New-AzureVMConfig -Name $vmName -InstanceSize $size -ImageName $imageName -ErrorAction Stop
    }

    $vmConfig | Set-AzureSubnet -SubnetNames $subnetNames -ErrorAction Stop | Out-Null
	
    if ($dataDisks -ne $null)
    {
	    for($i=0; $i -lt $dataDisks.Count; $i++)
	    {
	  	    $fields = $dataDisks[$i].Split(':')
		    $dataDiskLabel = [string] $fields[0]
	  	    $dataDiskSize = [string] $fields[1]
	  	    Write-Host ("Adding disk {0} with size {1}" -f $dataDiskLabel, $dataDiskSize)	
		
		    #Add Data Disk to the newly created VM
		    $vmConfig | Add-AzureDataDisk -CreateNew -DiskSizeInGB $dataDiskSize -DiskLabel $dataDiskLabel -LUN $i -ErrorAction Stop | Out-Null
	    }
    }
	
	if($windowsDomain)
	{	
		$vmConfig | Add-AzureProvisioningConfig -WindowsDomain -Password $password -AdminUserName $adminUserName -JoinDomain $domainjoin -Domain $domain -DomainPassword $domainPassword -DomainUserName $domainUserName -ErrorAction Stop | Out-Null
	}
	else
	{
		$vmConfig | Add-AzureProvisioningConfig -Windows -Password $password -AdminUserName $adminUserName -ErrorAction Stop | Out-Null
	}	

    if($newService)
    {
        if($windowsDomain)
        {
            $dns = New-AzureDns -Name "DNS" -IPAddress $dnsIP
    		New-AzureVM -ServiceName $serviceName -AffinityGroup $affinityGroup -VNetName $vnetName -DnsSettings $dns -VMs $vmConfig -WaitForBoot -Verbose -ErrorAction Stop
        }
        else
        {
    		New-AzureVM -ServiceName $serviceName -AffinityGroup $affinityGroup -VNetName $vnetName -VMs $vmConfig -WaitForBoot -Verbose -ErrorAction Stop
        }
    }
    else
    {
		New-AzureVM -ServiceName $serviceName -VMs $vmConfig -WaitForBoot -Verbose -ErrorAction Stop
    }

    InstallWinRMCertificateForVM $serviceName $vmName
    # Download remote desktop file to working directory (just in case)
    Get-AzureRemoteDesktopFile -ServiceName $serviceName -Name $vmName -LocalPath "$scriptFolder\$vmName.rdp" -ErrorAction Stop | Out-Null
    
    Write-Host "Pausing for Services to Start"
    Start-Sleep 60 

}

function RunWSManScriptBlock (
$serviceName,
$vmName,
$userName,
$password,
$credSSP = $false,
$argumentList,
$scriptBlock
)
{
    $uri = Get-AzureWinRMUri -ServiceName $serviceName -Name $vmName -ErrorAction Stop
    $credential = New-Object System.Management.Automation.PSCredential($userName, $(ConvertTo-SecureString $password -AsPlainText -Force))
    if($credSSP)
    {
        Invoke-Command `
            -ConnectionUri $uri.ToString() `
            -Credential $credential `
            -EnableNetworkAccess `
            -Authentication Credssp `
            -ArgumentList $argumentList `
            -ScriptBlock $scriptBlock `
            -ErrorAction Stop
    }
    else
    {
        Invoke-Command `
            -ConnectionUri $uri.ToString() `
            -Credential $credential `
            -ArgumentList $argumentList `
            -ScriptBlock $scriptBlock `
            -ErrorAction Stop
    }
}

function RunWSManScriptFile (
$serviceName,
$vmName,
$userName,
$password,
$credSSP = $false,
$argumentList,
$scriptFile
)
{
    $uri = Get-AzureWinRMUri -ServiceName $serviceName -Name $vmName
    $credential = New-Object System.Management.Automation.PSCredential($userName, $(ConvertTo-SecureString $password -AsPlainText -Force))
    if($credSSP)
    {
        Invoke-Command `
            -ConnectionUri $uri.ToString() `
            -Credential $credential `
            -EnableNetworkAccess `
            -Authentication Credssp `
            -ArgumentList $argumentList `
            -FilePath $scriptFile `
            -ErrorAction Stop
    }
    else
    {
        Invoke-Command `
            -ConnectionUri $uri.ToString() `
            -Credential $credential `
            -ArgumentList $argumentList `
            -FilePath $scriptFile `
            -ErrorAction Stop
    }
}

################## end Functions ##############################


################## Script execution ###########################

#Call Deploy
Deploy

################## end Script execution #######################

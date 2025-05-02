Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction "SilentlyContinue" | Out-Null
        
    $ErrorActionPreference = "Stop"
    
    if ($Verbose -eq $null) { $Verbose = $false }

# Ensure the SPTimerService is started on each Application Server
foreach ($server in (get-spserver | Where {$_.Role -eq "Application"}) )
{
	Write-Host "Starting SPTimerService on each Application Server"
	$server.Name
	$service = Get-WmiObject -computer $server.Name Win32_Service -Filter "Name='SPTimerV4'"
	$service.InvokeMethod('StopService',$Null)
	start-sleep -s 5
	$service.InvokeMethod('StartService',$Null)
	start-sleep -s 5
	$service.State
}

#Add Web Applications
[xml]$w = get-content WebApplications.xml

foreach ($WebApplication in $w.Setup.WebApplications.WebApplication)
{
		$WebAppName = $WebApplication.getAttribute("Name")
		$AppPool = $WebApplication.getAttribute("ApplicationPool")
		$AppPoolAcct = $WebApplication.getAttribute("ApplicationPoolAccount")
		$AuthMethod = $WebApplication.getAttribute("AuthenticationMethod")
		$DatabaseServer = $WebApplication.getAttribute("DatabaseServer")
		$DatabaseName = $WebApplication.getAttribute("DatabaseName")
		$DatabaseCred = $WebApplication.getAttribute("DatabaseCredentials")
		$HostHeader = $WebApplication.getAttribute("HostHeader")
		$Port = $WebApplication.getAttribute("Port")
		$URL = $WebApplication.getAttribute("Url")
		$Desc = $WebApplication.getAttribute("Description")
		$SSL = $WebApplication.getAttribute("SSL")

		$WebAppUrlPath = $URL
		Echo "-URL $WebAppUrlPath -Name $WebAppName -ApplicationPool $AppPool -ApplicationPoolAccount $AppPoolAcct -AuthenticationMethod $AuthMethod -DatabaseServer $DatabaseServer -DatabaseName $DatabaseName -AuthenticationProvider $ap -SecureSocketsLayer"

		echo " "
		echo "Checking for Web Application: $WebAppName"
		$TestWebApp = Get-SPWebApplication $URL -ErrorVariable err -ErrorAction SilentlyContinue 
		if ($err) 
		{ 
			if ($SSL -eq "true")
			{
	   			echo "Creating Web Application: $WebAppName"
				$AppPoolAccount = Get-SPManagedAccount $AppPoolAcct
				$ap = New-SPAuthenticationProvider -UseWindowsIntegratedAuthentication -DisableKerberos
				new-spwebapplication -Name $WebAppName -ApplicationPool $AppPool -ApplicationPoolAccount $AppPoolAccount -AuthenticationMethod $AuthMethod -AuthenticationProvider $ap -HostHeader $HostHeader -DatabaseName $DatabaseName -Port $Port -Url $URL -SecureSocketsLayer
				echo "Web Application Created"
				echo " "
				echo " "
			}
			else
			{
				echo "Creating Web Application: $WebAppName"
				$AppPoolAccount = Get-SPManagedAccount $AppPoolAcct
				$ap = New-SPAuthenticationProvider -UseWindowsIntegratedAuthentication -DisableKerberos
				new-spwebapplication -Name $WebAppName -ApplicationPool $AppPool -ApplicationPoolAccount $AppPoolAccount -AuthenticationMethod $AuthMethod -AuthenticationProvider $ap -HostHeader $HostHeader -DatabaseName $DatabaseName -Port $Port -Url $URL 
				echo "Web Application Created"
				echo " "
				echo " "

			}
		} 
		else 
		{
			echo "Web Application: $WebAppName already exists"
			echo " "
			echo " "
		}

		#Create the managed paths


		foreach ($ManagedPath in $WebApplication.ManagedPath)
		{
			echo "Managed Path: $ManagedPath" 
			if ($ManagedPath -ne $null)
			{
				$RelativeURL = $ManagedPath.getAttribute("RelativeURL")
				$WebApp = $ManagedPath.getAttribute("WebApplication")
				$TestPath = $URL + "/" + $RelativeURL
				$TestManagedPath = Get-SPManagedPath $TestPath -WebApplication $WebApp -ErrorVariable err -ErrorAction SilentlyContinue
				try{

				#if ($err)
				#{
					echo "Creating a managed path"
					new-spmanagedpath $RelativeURL -WebApplication $WebApp
					echo "Managed Path created"
					echo " "
					echo " "
				#}
				}
				Catch
				{
					$ErrorMessage = $_.Exception.Message
					$FailedItem = $_.Exception.ItemName
					echo "Error Message: $ErrorMessage"
					echo " "
					echo "Failed Item: $FailedItem"
					echo " " 
				}
			} 
		}

}

#$w1 = Get-SPWebApplication –identity "Connect"
#$w2 = Get-SPWebApplication –identity "My"
#$w1.GrantAccessToProcessIdentity("kelsey-seybold\svcProdC2WTS")
#$w2.GrantAccessToProcessIdentity("kelsey-seybold\svcProdC2WTS")
#$w1.GrantAccessToProcessIdentity("kelsey-seybold\svcProdExcel")
#$w2.GrantAccessToProcessIdentity("kelsey-seybold\svcProdExcel")
#$w1.GrantAccessToProcessIdentity("kelsey-seybold\svcProdPerfPt")
#$w2.GrantAccessToProcessIdentity("kelsey-seybold\svcProdPerfPt")
#$w1.GrantAccessToProcessIdentity("kelsey-seybold\svcProdPPivot")
#$w2.GrantAccessToProcessIdentity("kelsey-seybold\svcProdPPivot")
#$w1.GrantAccessToProcessIdentity("kelsey-seybold\svcProdSSRS")
#$w2.GrantAccessToProcessIdentity("kelsey-seybold\svcProdSSRS")

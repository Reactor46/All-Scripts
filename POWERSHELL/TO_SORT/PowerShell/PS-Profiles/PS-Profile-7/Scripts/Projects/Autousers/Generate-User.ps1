param ( 
	$FirstN, 
	$MI,
	$LastN, 
	$Title, 
	$Office, 
	$Department, 
	$Template, 
	$ScriptPath,
	[bool]$HomeDirectory=$True, 
	[bool]$Disabled=$True, 
	[bool]$ActiveSyncEnabled=$False 
	)
#User Variables
$alias = $FirstN.toLower().substring(0,1)+$LastN.tolower()
$aliaswithMI = $FirstN.toLower().substring(0,1)+$MI.tolower()+$LastN.tolower()
$Password = ConvertTo-SecureString 'mouse99!' -AsPlainText -Force
Write-Host "Creating New User: $FirstN $MI $LastN"
Function Set-HomeDirectory($alias, $Department, $Office)
	{
	switch ($Department) 
		{
		BH 
			{ 
			$FileServer = "misfs1"
			$NewFolder = Join-Path "\\$FileServer\d`$\User Shares\" $alias
			$LocalPath = Join-Path "D:\User Shares\" $alias
			}
		IDD 
			{
			if ( $office -match "Plano" )
				{
				$NoHome = $True
				$Stop = $True
				}
			else 
				{
				$FileServer = misfs1
				$NewFolder = Join-Path "\\$($FileServer)\d`$\IDD Users\" $alias
				$LocalPath = Join-Path "D:\IDD Users\" $alias
				}
			} 
		ECI 
			{ 
			$FileServer = "misfs1"
			$NewFolder = Join-Path "\\$($FileServer)\d`$\ECI USERS\" $alias
			$LocalPath = Join-Path "D:\ECI USERS\" $alias
			}
		Hotline { $NoHome = $True }
		}
		if ( !$NoHome )
			{
			$SharePath = "\\$($FileServer)\$($alias)"
			Write-Host "Creating Share: $SharePath" -ForegroundColor Yellow
			New-Item $NewFolder -Type Directory | Out-Null
			#Add-NTFSAccess -Path $NewFolder -Account "CCMHMR\$alias" -AccessRights Modify
			New-SMBShare –Name $alias –Path $LocalPath -FullAccess Everyone -CimSession $FileServer	| Out-Null
			Write-Host "Setting N Drive to $SharePath" -ForegroundColor Yellow		
			Set-AdUser -Identity $alias -HomeDirectory $SharePath -HomeDrive "N" -Server DC01
			}
		else
			{
			Write-Host "No HomeDirectory for $alias" -ForegroundColor Yellow
			if ( $Stop )
				{
				Write-Host "$alias is IDD Plano. You will have to go set up HomeDirectory in the Active Directory Users and Computers Console"
				}
			}
	}

Write-Host "Checking if username already exists"

If ( ( Find-ADuser $alias -Server DC01 ) -eq $null ) 
	{
	Write-Host "Creating $alias" -ForegroundColor Yellow
	$FullN = "$FirstN $LastN"
	$principal = $alias+"@ccmhmr.local"
	$email = $alias+"@lifepathsystems.org"
	}
elseif ( ( Find-Aduser $aliaswithMI -Server DC01 ) -eq $null )
	{
	Write-Host "$alias already exists."
	$alias = $aliaswithMI
	Write-Host "Creating $alias" -ForegroundColor Yellow
	$FullN = "$($FirstN) $($MI). $($LastN)"
	$principal = $alias+"@ccmhmr.local"
	$email = $alias+"@lifepathsystems.org"
	}
else 
	{
	Write-Host "Both $alias and $aliaswithMI taken. Cancelling." -ForegroundColor Red
	$Cancel = $True
	}

If ( !$Cancel )
	{
	New-Mailbox -UserPrincipalName $principal -PrimarySmtpAddress $email -alias $alias -Name $fulln -password $Password -FirstName $firstn -LastName $lastn -DisplayName $fulln -ResetPasswordOnNextLogon $true -erroraction stop | Out-Null
	Write-Host "Setting ActiveSync and OWA Access" -ForegroundColor Yellow
	Set-CASMailbox -Identity $alias -ActiveSyncEnabled $ActiveSyncEnabled -owaenabled $false
	Start-Sleep 10
	if ( $Disabled )
		{
		Write-Host "Disabling User and Hiding From Address Book" -ForegroundColor Yellow
		Set-Mailbox -Identity $alias -HiddenFromAddressListsEnabled $True
		Set-ADUser -Identity $alias -Enabled $False -Server DC01
		}
	Set-ADUser $alias -Department $Department -Office $Office -Title $Title -Description $Title -ScriptPath $ScriptPath -Server dc01
	Set-HomeDirectory -alias $alias -Department $Department -Office $Office
	Write-Host "Adding Group Memberships" -ForegroundColor Yellow
	$groups = (Get-ADUser $Template -Properties memberof).memberof
	$groups | Get-ADGroup -Server DC01 | Add-ADGroupMember -Members $alias -Server dc01
	}
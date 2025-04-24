param ( $FirstN, $MI,$LastN, $Title, $Office, $Template )
#User Variables
$FullN = "$FirstN $LastN"
$alias = $FirstN.toLower().substring(0,1)+$LastN.tolower()
$email = $alias+"@lifepathsystems.org"
$Password = ConvertTo-SecureString 'mouse99!' -AsPlainText -Force


If ( ( Find-ADuser $alias ) -eq $null ) 
	{
	New-mailbox -UserPrincipalName $email -alias $alias -Name $fulln -password $Password -FirstName $firstn -LastName $lastn -DisplayName $fulln -ResetPasswordOnNextLogon $true -erroraction stop
	Set-CASMailbox -Identity $alias -ActiveSyncEnabled $false -owaenabled $false
	Sleep 10
	Set-ADUser $alias -Office $Office -Title $Title -Server dc01
	$groups = (Get-ADUser $Template -Properties memberof).memberof
	$groups | Get-ADGroup | Add-ADGroupMember -Members $alias -Server dc01
	}

param([Parameter(Mandatory=$true)][string]$Filter)
Function Select-Computer
	{
	$Computers = Find-ADComputer $filter
	if ( $Computers.Count -ne '' )
		{
		$menu = @()
		for ($i=1;$i -le $Computers.count; $i++) 
			{
			Write-Host "$i. $($Computers[$i-1].name)" -ForegroundColor Cyan
		    $Computer = New-Object System.Object
			$Computer | Add-Member -MemberType NoteProperty -Name 'Index' -Value $i
			$Computer | Add-Member -MemberType NoteProperty -Name 'Name' -Value $($Computers[$i-1].name)
			$menu += $Computer
		    }
		$selection = Read-Host 'Selection'
		try
			{
			[int]$num = $selection
			}
		catch
			{
			Write-Host "Not a Valid Entry" -ForegroundColor Red
			exit
			}
		if ( $num -lt 1 -or $num -gt $menu.Count )
			{
			Write-Host "Not a Valid Entry" -ForegroundColor Red
			exit
			}
		$selection = $menu | ? { $_.Index -eq $selection }
		Return $Selection
		}
	elseif ( $Computers -eq $null )
		{
		Write-Host "No Match Found" -ForegroundColor Red
		exit
		}
	else
		{
		Return $Computers
		}
	}

if ( ( $Filter -as [ipaddress] ) -and ( $Filter -match "[.]" ) )
	{
	$Computer = $Filter
	}
else
	{
	$Computer = (Select-Computer).name
	}

Write-Host "Offering Remote Assistance to:" $Computer -ForegroundColor Yellow
msra /offerRA $Computer
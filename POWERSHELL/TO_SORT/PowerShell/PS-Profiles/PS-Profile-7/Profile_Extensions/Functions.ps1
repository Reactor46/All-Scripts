Function Wake-BoardCloset
	{
	Send-WOL 48:4D:7E:CF:41:4A
	}

Function Get-VerseoftheDay
	{
	$Scripture = (Get-BibleVerse -VerseOfTheDay -Type Json | ConvertFrom-Json)[0] 
	$book = $Scripture.bookname 
	$chap = $Scripture.chapter 
	$verse = $Scripture.verse 
	$ref = $book+" "+$chap+":"+$verse
	$txt = ('"' + $Scripture.text + '"').replace(' "','"') -Replace "  ", " "
	$space = " "*($txt.length - $ref.length)
        $border = ("="*$($txt.length)).substring(0, [System.Math]::Min(240, $txt.length))
        Write-Host $border -ForegroundColor DarkGray
	Write-Host $txt -ForegroundColor White
	Write-Host (" "*($txt.length))
	Write-Host $space$ref -ForegroundColor Yellow
        Write-Host $border -ForegroundColor DarkGray
	}

Function Color-LS
    {
    $regex_opts = ([System.Text.RegularExpressions.RegexOptions]::IgnoreCase `
          -bor [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $fore = $Host.UI.RawUI.ForegroundColor
    $compressed = New-Object System.Text.RegularExpressions.Regex(
          '\.(zip|tar|gz|rar|jar|war)$', $regex_opts)
    $executable = New-Object System.Text.RegularExpressions.Regex(
          '\.(exe|bat|cmd|py|pl|ps1|psm1|vbs|rb|reg)$', $regex_opts)
    $text_files = New-Object System.Text.RegularExpressions.Regex(
          '\.(txt|cfg|conf|ini|csv|log|xml|java|c|cpp|cs)$', $regex_opts)

    Invoke-Expression ("Get-ChildItem $args") | ForEach-Object {
        if ($_.GetType().Name -eq 'DirectoryInfo') 
            {
            $Host.UI.RawUI.ForegroundColor = 'Cyan'
            echo $_
            $Host.UI.RawUI.ForegroundColor = $fore
            }
        elseif ($compressed.IsMatch($_.Name)) 
            {
            $Host.UI.RawUI.ForegroundColor = 'DarkGreen'
            echo $_
            $Host.UI.RawUI.ForegroundColor = $fore
            }
        elseif ($executable.IsMatch($_.Name))
            {
            $Host.UI.RawUI.ForegroundColor = 'Yellow'
            echo $_
            $Host.UI.RawUI.ForegroundColor = $fore
            }
        elseif ($text_files.IsMatch($_.Name))
            {
            $Host.UI.RawUI.ForegroundColor = 'White'
            echo $_
            $Host.UI.RawUI.ForegroundColor = $fore
            Write-Host "_"
	    }
        else
            {
            echo $_
            }
        }
    }

Function Enter-SSHSession($Hostname,$User,$Password,$Port=22)
	{
	if ( $user -eq $null )
		{
		putty -ssh $hostname $Port
		}
	elseif ( $Password -eq $null )
		{
		putty -ssh $user@$hostname $Port
		}
	else
		{
		putty -ssh $user@$hostname -pw $Password $Port
		}
	}
Function Show-Desktop
	{
	(New-Object -ComObject shell.application).toggleDesktop()
	}	

Function Get-NewDCs
	{
	get-addomaincontroller -Filter { OperatingSystem -notlike "Windows Server 2003" -and OperatingSystem -notlike "Windows Server® 2008 Standard" }
	}

Function Sync-Files
	{
	Get-ChildItem $Docs\"Remote Assistance Logs" | Remove-Item -Confirm:$False
	robocopy C:\Users\wreeves\Documents \\misfs1\wreeves\Documents /MIR /FFT /Z /XA:H /W:5
	}

Function Get-UserVariable ($Name = '*')
	{
	# these variables may exist in certain environments (like ISE, or after use of foreach)
	$special = 'ps','psise','psunsupportedconsoleapplications', 'foreach', 'profile'
 	$ps = [PowerShell]::Create()
	$null = $ps.AddScript('$null=$host;Get-Variable') 
	$reserved = $ps.Invoke() |
 	Select-Object -ExpandProperty Name
	$ps.Runspace.Close()
	$ps.Dispose()
	Get-Variable -Scope Global |
	Where-Object Name -like $Name |
	Where-Object { $reserved -notcontains $_.Name } |
	Where-Object { $special -notcontains $_.Name } |
	Where-Object Name
	}

Function Get-FolderSize
	{
	param(
	[parameter(ValueFromPipelineByPropertyName)]
	$FullName = (pwd),
	[ValidateSet('TB', 'GB', 'MB', 'KB')]
	$Unit = "GB"
	)
	$munit = "1$($unit)"
	$item = Get-Item $Fullname
	$colitems = (Get-ChildItem $Fullname -R | Measure-Object -property length -sum)
	$item | Add-Member -MemberType NoteProperty -Name Size -Value ("{0:N2}" -f ($colItems.sum / $munit) + " $($unit)")
	$item | select Mode, LastWriteTime, Size, Name
	}

Function Get-WeatherPicture($City="75069") 
	{
	if ($City -eq $null)
		{
	        $City = Read-Host "Specify City"
	    	}
	(Invoke-WebRequest "http://wttr.in/$City" -UserAgent curl).content -split "`n"
	}

Function Suspend-Computer
	{
	# 1. Define the power state you wish to set, from the
	#    System.Windows.Forms.PowerState enumeration.
	$PowerState = [System.Windows.Forms.PowerState]::Suspend;

	# 2. Choose whether or not to force the power state
	$Force = $false;

	# 3. Choose whether or not to disable wake capabilities
	$DisableWake = $false;

	# Set the power state
	[System.Windows.Forms.Application]::SetSuspendState($PowerState, $Force, $DisableWake);
	}

Function Get-WeekOfYear
	{
	Get-Date -UFormat %V
	}

Function Send-FileHome($FileName,$Destination)
	{
	kscp -P 2200 $filename wayne@darthwayne.duckdns.org:/home/wayne/$destination
	}

Function Get-IPInfo
    {
    param($ComputerName="localhost")
    get-wmiobject win32_networkadapterconfiguration -computername $computername | 
        Where-Object { $_.ipaddress -ne $null } | 
            Select-Object Description, IPaddress, IPSubnet, DefaultIPGateway, DNSServerSearchOrder, MacAddress
    }

Function Search-ArchWiki
    {
    param(
        [Parameter(Mandatory=$True)]
        [string]$Query
    )
    if ( $Query -ne $null )
        {
        $Query = $Query.Replace(" ","+")
        $URL = "https://wiki.archlinux.org/index.php?search=$($Query)&title=Special%3ASearch&profile=default&fulltext=1"
        }
    else
        {
        $URL = "https://wiki.archlinux.org"
        }
    Start-Process $URL
    }

Function Search-MISWiki
    {
    param(
            [Parameter(Mandatory=$False)]
            [string]$Query
         )
    if ( $Query -ne $null )
        {
        $Query = $Query.Replace(" ","+")
        $URL = "http://miswiki/doku.php?do=search&id=$($Query)"
        }
    else
        {
        $URL = "http://miswiki"
        }
    Start-Process $URL
    }

Function Reload-misModules
    {
    get-module mis* | remove-module | import-module
    }

Function Get-UserInfo
    {
    Param(
        [Parameter(
            Position=0
        )]
        [string]$Name,
        [Parameter(
            Position=1
        )]
        [string]$Title,
        [Parameter(
            Position=2
        )]
        [string]$Department,
        [Parameter(
            Position=3
        )]
        [string]$Office,

        [Parameter(
            Position=4
        )]
        [string]$Phone

    )
    If ( $Name )
        {
        $Name = "(|(name=*$($Name)*)(samaccountname=*$($Name)*))"
        }
    
    If ( $Department )
        {
        $Department = "(department=*$($Department)*)"
        }
    If ( $Title )
        {
        $Title = "(title=*$($Title)*)"
        }
    If ( $Office )
        {
        $Office = "(physicalDeliveryOfficeName=*$($Office)*)"
        }
    If ( $Phone )
        {
        $Phone = "(|(telephoneNumber=*$($Phone)*)(mobile=*$($Phone)*))"
        }
    $SearchString = "(&$($Name)$($Office)$($Department)$($Phone)$($Title)$($Phone))"
    Get-ADUser -ldapfilter $SearchString -Properties Office, Department, Title, MobilePhone, OfficePhone | ft Name, Title, Department, Office, OfficePhone, MobilePhone
    }


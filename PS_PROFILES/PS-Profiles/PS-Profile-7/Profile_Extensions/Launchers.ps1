Function Launch-Far
	{
	."C:\Program Files\Far Manager\Far.exe"
	}

Function Launch-Helpdesk
	{
	Start-Process "http://mishelp1:8080/WOListView.do?requestViewChanged=true&"
	}

Function Launch-CMBHS
	{
	chrome "https://cmbhs.dshs.state.tx.us/cmbhs/WebPages/Default.aspx"
 	}

Function Launch-Anasazi
	{
	#param(
	#[ValidateSet("Live", "Train", "Test")]
	#$Version = "Live"
	#)
	#Show-Desktop
	$staffid | clip
	$Launch = "$($Version) Anasazi Central"
	$runpath = "C:\Program Files (x86)\Citrix\ICA Client\SelfServicePlugin\SelfService.exe"
	.$runpath -showAppPicker
	#.$runpath -rmPrograms | Out-Null
	#.$runpath -poll | Out-Null
	#.$runpath -qlaunch $Launch
	#$ica = "https://misctrx1.ccmhmr.local/Citrix/Lifepath/resources/v2/WGVuQXBwNi41LkxpdmUgQW5hc2F6aSBDZW50cmFs/launch/ica"
	#.$runpath -launch -s lifepath-80143580 -CitrixID "lifepath-80143580@@XenApp6.5.Live Anasazi Central" -ica $ica -cmdline
	#.(get-childitem 
	}

Function Launch-AssetInventory
	{
	Invoke-Item -Path "\\Misfs1\svcreq$\Request for Services.mdb"
	}

Function Launch-Paylocity
	{
	Start-Process "https://login.paylocity.com/Escher/Escher_WebUI/views/login/login.aspx"
	}

Function Enter-Clonezilla
	{
	Enter-SSHSession -User administrator -Hostname 192.168.103.254 -Password ( Get-XMLPassword -Name Clonezilla -Type Password -AsPlainText $True )
	}

Function Start-ScreenConnect($computer)
	{
	Start-Process "https://assist.lifepathsystems.org/Host#Access/All%20Machines/$computer"
	}

Function Manage-EqualLogic
	{
	."$Scripts\Java\groupmgr.jnlp"
	}

Function Launch-Video($VideoID,[switch]$VLC)
	{
	$url = "https://www.youtube.com/embed/$VideoID"
        if ( $VLC )
            {
            ."C:\Program Files\VideoLAN\vlc\vlc.exe" $url
            }
        else
            {
            $ie = new-object -ComObject "InternetExplorer.Application"
            $ie.MenuBar = $False
            $ie.StatusBar = $False
            $ie.ToolBar = $False
            $ie.AddressBar = $False
            $ie.Top = 600
            $ie.Left = 1100
            $ie.Width = 480
            $ie.Height = 298
            $ie.Navigate($url)
            $ie.visible = $True
            }
	}

Function Search-YouTube($query, $results=10)
    {
    $params = @{
		type='video';
		q=$query;
                part='snippet';
                maxResults=$results;
                key='AIzaSyDTx7_0Xl09t4adX06h2pmJrsDVCqxtb1A'  
		}   

    $response = Invoke-RestMethod -Uri https://www.googleapis.com/youtube/v3/search -Body $params -Method Get
    for ( $i=1; $i -le $Response.items.count; $i++)
        {
        Write-Host "$i. $($response.items[$i-1].snippet.title)" -ForegroundColor Cyan
        }
    $Selection = Read-host "Selection"
    Launch-Video -VideoID ($response.items[$selection-1].id.videoId)
    }

Function Launch-CMBHS
	{
	Start-Process "https://cmbhs.dshs.state.tx.us/cmbhs/WebPages/Default.aspx"
	}

Function Launch-ChromeRemoteDesktop
	{
	."C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"  --profile-directory=Default --app-id=gbchcmhmhahfdphkhkmpfmihenigjmpp
	}

Function Launch-DellSupport
	{
	Start-Process "http://www.dell.com/support/incidents-online/us/en/19/contactus/dynamic"
	}

Function Launch-MISSharePoint
	{
	Start-Process "http://sharepoint/sites/mis/_layouts/15/start.aspx#/"
	}

Function Launch-SharePoint
	{
	Start-Process "http://sharepoint"
	}

Function SSH-Wayne
	{
        #kageant "C:\Users\wreeves\Documents\WayneKeyGen\WaynePriv.ppk"
	kitty -load wayne
	}

Function Select-AdministrativeTool
	{
	Param($Selection)
	Function Get-Selection
			{
			for ( $i=1;$i -le $Choices.count; $i++)
					{
					Write-Host "$i. $($Choices[$i-1].BaseName)" -ForegroundColor Cyan
					}
					$Selection = Read-Host "Selection"
					try
							{
							[int]$num = $Selection
							}
			Catch
							{
							Write-Host "Not a Valid Entry" -ForegroundColor Red
							}
					if ( $num -lt 1 -or $num -gt $Choices.count )
					{
					Write-Host "Not a Valid Entry" -ForegroundColor Red
					exit
					}
			$Selection = $Choices[$Selection-1]
			Return $Selection
			}
	$Choices = Get-ChildItem "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Administrative Tools"
	if ( $Selection -eq $null )
					{
					$Selection = Get-Selection
					}
	elseif ( $Selection -is [int] )
			{
			$Selection = $Choices[$Selection-1]
			}
	else
			{
			$Choices = $Choices | ? BaseName -match $Selection
			if ( $Choices.count -gt 1 )
					{
					$Selection = Get-Selection
					}
			else
					{
					$Selection = $Choices
					}
			}
	if ( $Selection.BaseName -eq "Remote Desktop Services" )
			{
			Write-Host "Launching $($Selection.BaseName)" -ForegroundColor Green
			ii $Selection.Fullname
			}
	else
			{
			Write-Host "Launching $($Selection.BaseName)" -ForegroundColor Green
			.($Selection.Fullname)
			}
	}

Function Get-ClosingTime
    {
    param(
        [switch]$ShutDown
        )
    $ClosingTime = Get-Date "$(Get-Date -Format "MM-dd-yy") 5:00PM"
    while ( (Get-Date) -le $ClosingTime )
            {
            $Date = Get-Date 
            $Countdown = $ClosingTime - $Date
            Write-Host "`r$(Get-Date $Date -Format "hh:mm:ss") Countdown: -$($Countdown.Hours):$($Countdown.Minutes):$($Countdown.Seconds):$($Countdown.milliseconds)   " -ForegroundColor Yellow -NoNewLine
            }
    if ( $ShutDown )
        {   
        sd -f
        }
    else
        {
        Write-Host "Go Home!" -ForegroundColor Green
        }
    }

Function Launch-VisualTime
    {
    ."\\misfs1\Global\Meeting Room Scheduler\visualtime.exe" -URL=http://missched1/visualtime
    }


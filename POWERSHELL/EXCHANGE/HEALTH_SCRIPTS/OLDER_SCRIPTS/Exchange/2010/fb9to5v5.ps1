$UserName = "admin@domain.com"
$Password = "password1"
$secpassword = ConvertTo-SecureString $Password -AsPlainText -Force
$adminCredential = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $UserName,$secpassword   

If(Get-PSSession | where-object {$_.ConfigurationName -eq "Microsoft.Exchange"}){
	write-host "Session Exists"
}
else{
	$rpRemotePowershell = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -credential $adminCredential  -Authentication Basic -AllowRedirection 
	$importresults = Import-PSSession $rpRemotePowershell 
} 


$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind
$service.Credentials = New-Object System.Net.NetworkCredential($username,$password)
$service.AutodiscoverUrl($UserName ,{$true})

$mbHash = @{ }

$tmValHash = @{ }
$tidx = 0
for($vsStartTime=[DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 0:00"));$vsStartTime -lt [DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 0:00")).AddDays(1);$vsStartTime = $vsStartTime.AddMinutes(30)){
	$tmValHash.add($vsStartTime.ToString("HH:mm"),$tidx)	
	$tidx++
}

get-mailbox -ResultSize unlimited | foreach-object{
	if ($mbHash.ContainsKey($_.WindowsEmailAddress.ToString()) -eq $false){
		$mbHash.Add($_.WindowsEmailAddress.ToString(),$_.DisplayName)
	}
}
$Attendeesbatch = [activator]::createinstance(([type]'System.Collections.Generic.List`1').makegenerictype([Microsoft.Exchange.WebServices.Data.AttendeeInfo]))

$StartTime = [DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 0:00"))
$EndTime = $StartTime.AddDays(1)


$displayStartTime =  [DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 08:30"))
$displayEndTime = [DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 17:30"))

$drDuration = new-object Microsoft.Exchange.WebServices.Data.TimeWindow($StartTime,$EndTime)
$AvailabilityOptions = new-object Microsoft.Exchange.WebServices.Data.AvailabilityOptions
$AvailabilityOptions.RequestedFreeBusyView = [Microsoft.Exchange.WebServices.Data.FreeBusyViewType]::DetailedMerged

$batchsize = 100
$bcount = 0
$bresult = @()
if ($mbHash.Count -ne 0){
	$mbHash.GetEnumerator() | Sort Value | foreach-object {
		if ($bcount -ne $batchsize){
			$Attendee1 = new-object Microsoft.Exchange.WebServices.Data.AttendeeInfo($_.Key)
			$Attendeesbatch.add($Attendee1)
			$bcount++
		}
		else{
			$availresponse = $service.GetUserAvailability($Attendeesbatch,$drDuration,[Microsoft.Exchange.WebServices.Data.AvailabilityData]::FreeBusy,$AvailabilityOptions)
			foreach($avail in $availresponse.AttendeesAvailability.OverallResult){$bresult += $avail}
			$Attendeesbatch = [activator]::createinstance(([type]'System.Collections.Generic.List`1').makegenerictype([Microsoft.Exchange.WebServices.Data.AttendeeInfo]))
			$bcount = 0
			$Attendee1 = new-object Microsoft.Exchange.WebServices.Data.AttendeeInfo($_.Key)
			$Attendeesbatch.add($Attendee1)
			$bcount++
		}
	}
}
$availresponse = $service.GetUserAvailability($Attendeesbatch,$drDuration,[Microsoft.Exchange.WebServices.Data.AvailabilityData]::FreeBusy,$AvailabilityOptions)
$usrIdx = 0
$frow = $true
foreach($res in $availresponse.AttendeesAvailability){
      if ($frow -eq $true){
		$fbBoard = $fbBoard + "<table><tr bgcolor=`"#95aedc`">" +"`r`n"
		$fbBoard = $fbBoard + "<td align=`"center`" style=`"width=200;`" ><b>User</b></td>" +"`r`n"
		for($stime = $displayStartTime;$stime -lt $displayEndTime;$stime = $stime.AddMinutes(30)){
			$fbBoard = $fbBoard + "<td align=`"center`" style=`"width=50;`" ><b>" + $stime.ToString("HH:mm") + "</b></td>" +"`r`n"
		}
		$fbBoard = $fbBoard + "</tr>" + "`r`n"
		$frow = $false
	}
	for($stime = $displayStartTime;$stime -lt $displayEndTime;$stime = $stime.AddMinutes(30)){
		if ($stime -eq $displayStartTime){
			$fbBoard = $fbBoard + "<td bgcolor=`"#CFECEC`"><b>" + $mbHash[$Attendeesbatch[$usrIdx].SmtpAddress] + "</b></td>"  + "`r`n"
		}
		$title = "title="
		if ($res.MergedFreeBusyStatus[$tmValHash[$stime.ToString("HH:mm")]] -ne $null){
			$gdet = $false
			$FbValu = $res.MergedFreeBusyStatus[$tmValHash[$stime.ToString("HH:mm")]]
			switch($FbValu.ToString()){
				"Free" {$bgColour = "bgcolor=`"#41A317`""}
				"Tentative" {$bgColour = "bgcolor=`"#52F3FF`""
					     $gdet = $true
					}
				"Busy" {$bgColour = "bgcolor=`"#153E7E`""
					     $gdet = $true
					}
				"OOF" {$bgColour = "bgcolor=`"#4E387E`""
					     $gdet = $true
					}
				"NoData" {$bgColour = "bgcolor=`"#98AFC7`""}
				"N/A" {$bgColour = "bgcolor=`"#98AFC7`""}		
			}
			if ($gdet -eq $true){
				if ($res.CalendarEvents.Count -ne 0){
					for($ci=0;$ci -lt $res.CalendarEvents.Count;$ci++){
						if ($res.CalendarEvents[$ci].StartTime -ge $stime -band $stime -le $res.CalendarEvents[$ci].EndTime){				
							if($res.CalendarEvents[$ci].Details.IsPrivate -eq $False){
								$subject = ""
								$location = ""
								if ($res.CalendarEvents[$ci].Details.Subject -ne $null){
									$subject = $res.CalendarEvents[$ci].Details.Subject.ToString()
								}
								if ($res.CalendarEvents[$ci].Details.Location -ne $null){
									$location = $res.CalendarEvents[$ci].Details.Location.ToString()
								}
								$title = $title + "`"" + $subject + " " + $location + "`" "
							}
						}
					}
				}
			}
			
		}
		else{
			$bgColour = "bgcolor=`"#98AFC7`""
		}
		if($title -ne "title="){
			$fbBoard = $fbBoard + "<td " + $bgColour + " " + $title + "></td>"  + "`r`n"
		}
		else{
			$fbBoard = $fbBoard + "<td " + $bgColour + "></td>"  + "`r`n"
		}

	}
	$fbBoard = $fbBoard + "</tr>"  + "`r`n"
	$usrIdx++
}
$fbBoard = $fbBoard + "</table>"  + "  " 
$fbBoard | out-file "c:\fbboard.htm"
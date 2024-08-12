[void][Reflection.Assembly]::LoadFile("C:\temp\EWSUtil.dll")

$casUrl = "https://servername/ews/exchange.asmx"
$mbHash = @{ }

get-mailbox -ResultSize unlimited | foreach-object{
	if ($mbHash.ContainsKey($_.WindowsEmailAddress.ToString()) -eq $false){
		$mbHash.Add($_.WindowsEmailAddress.ToString(),$_.DisplayName)
	}
}
$mbs = @()

$ewc = new-object EWSUtil.EWSConnection($mbMailboxEmail,$false, "username", "password", "domain",$casUrl)
$drDuration = new-object EWSUtil.EWS.Duration
$drDuration.StartTime = [DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 08:30"))
$drDuration.EndTime = [DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 17:30"))

$batchsize = 100
$bcount = 0
$bresult = @()
if ($mbHash.Count -ne 0){
	foreach($key in $mbHash.keys){
		if ($bcount -ne $batchsize){
			$mbs += $key
			$bcount++
		}
		else{
			$bresult += $ewc.GetAvailiblity($mbs, $drDuration, 30)
			$mbs = @()	
			$bcount = 0
			$mbs += $key
			$bcount++
		}
	}
}
$bresult += $ewc.GetAvailiblity($mbs, $drDuration, 30)
$frow = $true
foreach($fbents in $bresult){
	foreach($key in $fbents.keys){
		if ($frow -eq $true){
			$fbBoard = $fbBoard + "<table><tr bgcolor=`"#95aedc`">" +"`r`n"
			$fbBoard = $fbBoard + "<td align=`"center`" style=`"width=200;`" ><b>User</b></td>" +"`r`n"
			for($stime = $drDuration.StartTime;$stime -lt $drDuration.EndTime;$stime = $stime.AddMinutes(30)){
				$fbBoard = $fbBoard + "<td align=`"center`" style=`"width=50;`" ><b>" + $stime.ToString("HH:mm") + "</b></td>" +"`r`n"
			}
			$fbBoard = $fbBoard + "</tr>" + "`r`n"
			$frow = $false
		}
		for($stime = $drDuration.StartTime;$stime -lt $drDuration.EndTime;$stime = $stime.AddMinutes(30)){
			$valuehash = $fbents[$key]
			if ($stime -eq $drDuration.StartTime){
				$fbBoard = $fbBoard + "<td bgcolor=`"#CFECEC`"><b>" + $mbHash[$valuehash[$stime.ToString("HH:mm")].MailboxEmailAddress.ToString()] + "</b></td>"  + "`r`n"
			}
			switch($valuehash[$stime.ToString("HH:mm")].FBStatus.ToString()){
				"0" {$bgColour = "bgcolor=`"#41A317`""}
				"1" {$bgColour = "bgcolor=`"#52F3FF`""}
				"2" {$bgColour = "bgcolor=`"#153E7E`""}
				"3" {$bgColour = "bgcolor=`"#4E387E`""}
				"4" {$bgColour = "bgcolor=`"#98AFC7`""}
				"N/A" {$bgColour = "bgcolor=`"#98AFC7`""}		
			}
			$title = "title="
			if ($valuehash[$stime.ToString("HH:mm")].FBSubject -ne $null){
				if ($valuehash[$stime.ToString("HH:mm")].FBLocation -ne $null){
					$title =  $title + "`"" + $valuehash[$stime.ToString("HH:mm")].FBSubject.ToString() + " " + $valuehash[$stime.ToString("HH:mm")].FBLocation.ToString() + "`" "
				}
				else {
					$title =  $title + "`"" + $valuehash[$stime.ToString("HH:mm")].FBSubject.ToString() + "`" "
				}
			}
			else {
				if ($valuehash[$stime.ToString("HH:mm")].FBLocation -ne $null){
					$title =  $title + "`"" + $valuehash[$stime.ToString("HH:mm")].FBLocation.ToString() + "`" "
				}
			}
			if($title -ne "title="){
				$fbBoard = $fbBoard + "<td " + $bgColour + " " + $title + "></td>"  + "`r`n"
			}
			else{
				$fbBoard = $fbBoard + "<td " + $bgColour + "></td>"  + "`r`n"
			}

		}
		$fbBoard = $fbBoard + "</tr>"  + "`r`n"
		
	}
}
$fbBoard = $fbBoard + "</table>"  + "  " 
$fbBoard | out-file "c:\fbboard.htm"
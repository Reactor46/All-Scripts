$ServerName = $args[0] 

$ExcomCollection = @()
$MBHash = @{ }
$MBHistHash = @{ }

$HistoryDir = "c:\LogonHistory"

if (!(Test-Path -path $HistoryDir))
{
	New-Item $HistoryDir -type directory
	$frun = 1
}

$datetime = get-date
$fname = $script:HistoryDir + "\"
$fname = $fname  + $datetime.ToString("yyyyMMdd")  + $ServerName +  "-LogonHist.csv"

Import-Csv ($fname)  | %{ 
	$idvalue = $_.identity.ToString()
	$logonEvent = $_
	$ltime = [datetime]::Parse($_.LogonTime)
	if ($_.ClientIPAddress -eq $null){$agvalue = $_.identity.ToString() + $_.UserName + $ltime.ToString("hhmm").Substring(0,3)}
	else{$agvalue = $_.identity.ToString() + $_.UserName +  $_.ClientIPAddress}
	if ($MBHistHash.Containskey($agvalue) -eq $false){
		$MBHistHash.Add($agvalue,$_)
	}

}


get-mailbox -server $ServerName -ResultSize Unlimited | foreach-object{
	if ($MBHash.Containskey($_.LegacyExchangeDN.ToString()) -eq $false){
		$MBHash.add($_.LegacyExchangeDN.ToString(),$_)
	}
}
$LogonUnQ = @{ }

get-logonstatistics | foreach-object{
	$idvalue = $_.identity.ToString()
	$logonEvent = $_
	$ltime = $_.LogonTime
	if ($_.ClientIPAddress -eq $null){$agvalue = $_.identity.ToString() + $_.UserName + $ltime.ToString("hhmm").Substring(0,3)}
	else{$agvalue = $_.identity.ToString() + $_.UserName +  $_.ClientIPAddress}
	if ($idvalue  -ne $null){
		if ($LogonUnQ.Containskey($agvalue) -eq $false){
			if ($MBHash.Containskey($idvalue)){
				$LogonUnQ.Add($agvalue,$logonEvent)				
				if ($MBHistHash.Containskey($agvalue) -eq $false){
					$MBHistHash.Add($agvalue,$_)
				}
				else{
					$ts = New-timeSpan $MBHistHash[$agvalue].LogonTime $ltime  
					if ($ts.minutes -gt 5){
						$MBHistHash.Add($agvalue+$ltime,$_)

					}
				}
			}
	
		}
	}
}
foreach ($row in $MBHistHash.Values){
	$ExcomCollection += $row
}

$ExcomCollection | export-csv –encoding "unicode" -noTypeInformation $fname
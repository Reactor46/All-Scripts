$dtQueryDT = [system.DateTime]::Now.AddHours(-($args[1]))
$dtQueryDTf = [system.DateTime]::Now


$InternalNum = 0
$InternalSize = 0

$ExternalSentNum = 0
$ExternalSentSize = 0

$ExternalRecNum = 0
$ExternalRecSize = 0

$ServerName = $args[0]

$DomainHash = @{}
$msgIDArray = @{}


get-accepteddomain | ForEach-Object{
	if ($_.DomainType -eq "Authoritative"){
		$DomainHash.add($_.DomainName.SmtpDomain.ToString().ToLower(),1)
	}

}

Get-MessageTrackingLog -Server $ServerName -ResultSize Unlimited -Start $dtQueryDT -End $dtQueryDTf  | ForEach-Object{ 
	if ($_.EventID.ToString() -eq "SEND" -bor $_.EventID.ToString() -eq "RECEIVE"){
		foreach($recp in $_.recipients){
		if($recp.ToString() -ne ""){
			$unkey = $recp.ToString() + $_.Sender.ToString() + $_.MessageId.ToString()
			if ($msgIDArray.ContainsKey($unkey) -eq $false){
				$msgIDArray.Add($unkey,1)
				$recparray = $recp.split("@")
				$sndArray = $_.Sender.split("@")
				if ($_.Sender -ne ""){
				if ($DomainHash.ContainsKey($recparray[1])){
					if ($DomainHash.ContainsKey($sndArray[1])){
						$InternalNum = $InternalNum + 1
						$InternalSize = $InternalSize + $_.TotalBytes/1024
					}
					else{
	
						$ExternalRecNum = $ExternalRecNum + 1					
						$ExternalRecSize = $ExternalRecSize + $_.TotalBytes/1024
					}			
				}
				else{
					if ($DomainHash.ContainsKey($sndArray[1])){
						$ExternalSentNum = $ExternalSentNum + 1					
						$ExternalSentSize = $ExternalSentSize + $_.TotalBytes/1024	
					}				
				}
				}
			}
			
		}
		}     
	}

}

"Sent/Recieved Internally Number : " + $InternalNum 
"Sent/Recieved Internally Size : " + [math]::round($InternalSize/1024,2) 
"Externally Sent Number : " + $ExternalSentNum
"Externally Sent Size : " + [math]::round($ExternalSentSize/1024,2) 
"Externally Recieved Number : " + $ExternalRecNum
"Externally Recieved Size : " + [math]::round($ExternalRecSize/1024,2)


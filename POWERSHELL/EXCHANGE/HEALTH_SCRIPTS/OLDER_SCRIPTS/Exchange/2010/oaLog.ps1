[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 
$exFileName = new-object System.Windows.Forms.openFileDialog
$exFileName.ShowHelp = $true
$exFileName.ShowDialog()

$fname = $exFileName.FileName
$mbcombCollection = @()
$FldHash = @{}
$usHash = @{} 
$fieldsline = (Get-Content $fname)[3]
$fldarray = $fieldsline.Split(" ")
$fnum = -1
foreach ($fld in $fldarray){
	$FldHash.add($fld,$fnum)
	$fnum++
}

get-content $fname | Where-Object -FilterScript { $_ -ilike “*MSRPC*” } | %{ 
	$lnum ++
	if ($lnum -eq $rnma){ Write-Progress -Activity "Read Lines" -Status $lnum
		$rnma = $rnma + 1000
	}
	$linarr = $_.split(" ")
	$uid = $linarr[$FldHash["cs-username"]] + $linarr[$FldHash["c-ip"]]
	if ($linarr[$FldHash["cs-username"]].length -gt 2){
		if ($usHash.Containskey($uid) -eq $false){
			$usrobj = "" | select UserName,IpAddress
			$usrobj.UserName = $linarr[$FldHash["cs-username"]]
			$usrobj.IpAddress = $linarr[$FldHash["c-ip"]]
			$usHash.add($uid,$usrobj)
			$mbcombCollection += $usrobj
		
		}
	}
}

$mbcombCollection | export-csv –encoding "unicode" -noTypeInformation c:\oareport.csv
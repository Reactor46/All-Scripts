$ver = (get-host).Version.Major

if ([convert]::ToInt32($ver, 10) -lt 4){

	Write-Host "You need to update your powershell to version 4 before you can run this script!"
	$temp = Read-Host
} else {

	if (!(Test-Path variable:global:cred))
	{
		$cred = Get-Credential
	}

	$error
	$computersFile = Read-Host "Enter the path for the CSV source file"
	$Computers = Get-Content $computersFile

	$Reboot = Read-Host "Do you want to force an immediate reboot of successfully renamed machines? (y/n)"
	
	$bolValidated = $False
	
	do {
		
		if ($Reboot.ToLower().StartsWith("y")){
			$bolValidated = $True
			$Reboot = "y"
		} elseif ($Reboot.ToLower().StartsWith("n")){
			$bolValidated = $True
			$Reboot = "n"			
		} else {
            $Reboot = Read-Host "Please only enter y or n"
        }
	} while (!$bolValidated)
	
	Clear-Content -Path $computersFile

	foreach ($computer in $computers){
		
		$names = $computer.Split(',')
		$error.Clear()

		if ($names[0] -eq "Old Name"){
			
			$text = $names[0] + "," + $names[1] + "," + $names[2] + "," + $names[3] + "," + $names[4]
			Add-Content -Path $computersFile -Value $text
		} else {

			if ($names[3] -ne "Completed"){
			
				$oldName = $names[0]
				$newName = $names[1]
				$text = $oldName + "," + $newName

				if (Test-Connection -ComputerName $oldName -Count 1 -Quiet){
				
					$text = $text + ",Yes"

                    if ($Reboot.CompareTo("y")){
                    					
                        Rename-Computer -newName $newName -ComputerName $oldName -Restart -DomainCredential $cred 2> $Null
                    } else {
                        
                        Rename-Computer -newName $newName -ComputerName $oldName -DomainCredential $cred 2> $Null
                    }
					if ($error.length -gt 0){
						
						$text = $text + ",Failed," + $error
						Add-Content -Path $computersFile -Value $text
					} else {
						
						$text = $text + ",Completed"
						Add-Content -Path $computersFile -Value $text
					}
				} else {
					
					$text = $text + ",No,Failed"
					Add-Content -Path $computersFile -Value $text
				}
			}
		}
	}
}
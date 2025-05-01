$comps = Import-Csv P:\Inventory\IMEI\LIST1.csv
$credential = Get-Credential
foreach ( $comp in $comps )
	{
	$compname = $comp.computername
	$newname = $compname -Replace "avenues","bhmckinney"
	Write-Host "Attempting to rename $compname to $newname" -ForegroundColor Yellow
	If ( Test-Connection $compname -BufferSize 1 -Count 1 -Quiet )
		{
		Try
			{
			Rename-Computer -ComputerName $compname -NewName $newname -Confirm:$false -Restart -DomainCredential $credential
			}
		Catch
			{
			Write-Host "Failed to rename" -ForegroundColor Red
			}
		If ( !$Error )
			{
			Write-Host "Success" -ForegroundColor Green
			}
		}
	Else	
		{
		Write-Host "Failed to ping: $compname"
		}
	}
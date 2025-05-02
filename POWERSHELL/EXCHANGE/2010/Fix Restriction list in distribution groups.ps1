#################################################################################################
#			Author: Lokesh Agarwal
#			Reviewer: Vikas Sukhija
#			Modfied: 05/02/1015
#			Description: Fix Restriction list in distribution groups
#
#
#################################################################################################

Param(
  [string]$name
)

####################################Define Logs##############################

$date = get-date -format d
$date = $date.ToString().Replace(“/”, “-”)

$time = get-date -format t
$time = $time.ToString().Replace(":", "-")
$time = $time.ToString().Replace(" ", "")

$log1 = ".\Logs" + "\" + "ExistingR_" + $date + "_" + $time + "_.log"
$log2 = ".\Logs" + "\" + "NewR_" + $date + "_" + $time + "_.log"

##############################ADD Shell#####################################

If ((Get-PSSnapin | where {$_.Name -match "Microsoft.Exchange.Management.PowerShell.E2010"}) -eq $null)
{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
}

$name= read-host "Please enter group name that needs to be fixed"

$DLCurrent = Get-DistributionGroup -identity $name
$DLTerminateUser = @((Get-DistributionGroup -identity $name ).AcceptMessagesOnlyFrom | %{$_.distinguishedname})
$TermCount = $DLTerminateUser.Count
Write-host "Total restriction count $TermCount for $name"
Add-content $log1 "Total restriction count $TermCount for $name"
Add-content $log1 $DLTerminateUser

		while ($TermCount -ne 0)
		{
			if (!(Get-Mailbox $DLTerminateUser[$TermCount - 1]))
			{ if(!(Get-DistributionGroup -identity $DLTerminateUser[$TermCount - 1]))
					{
			$DLCurrent.AcceptMessagesOnlyFrom
			$DLCurrent.AcceptMessagesOnlyFrom -= $DLTerminateUser[$TermCount - 1]
			$DLCurrent.AcceptMessagesOnlyFrom
				      }
			}		  

		$TermCount = $TermCount - 1
		}

                 if($error -like "*couldn't be found*")
                   {
			$error.clear()
		   }
		$countitems= $DLCurrent.AcceptMessagesOnlyFrom
		$countit = $countitems.count
		Write-host "Total restriction count after removing Term users from $name is $countit"
		Add-content $log2 "Total restriction count after removing Term users from $name is $countit"
		Add-content $log2 $countitems

		timeout 20
Set-DistributionGroup -identity $name -AcceptMessagesOnlyFrom $DLCurrent.AcceptMessagesOnlyFrom		
		timeout 20

		$DLCurrent2 = Get-DistributionGroup -identity $name
		$DLTerminateUser2 = @((Get-DistributionGroup -identity $name ).AcceptMessagesOnlyFrom | %{$_.distinguishedname})

		$TermCount2 = $DLTerminateUser2.Count
                Write-host "Final restriction count after removing Term users from $name is $TermCount2"
		Add-content $log2 "Final restriction count after removing Term users from $name is $TermCount2"


##############################################################################################################
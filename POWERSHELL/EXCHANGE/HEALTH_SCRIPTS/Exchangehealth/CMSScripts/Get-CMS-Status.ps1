#############################################################################
#       Author: Mahesh Sharma
#       Reviewer: Vikas SUkhija      
#       Date: 06/10/2013
#	Modified:06/19/2013 - made it to run from any path
#       Description: ExChange Health Status
#############################################################################

########################### Add Exchange Shell###############################

If ((Get-PSSnapin | where {$_.Name -match "Exchange.Management"}) -eq $null)
{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin
}

############################Define Variables#################################

$report = "\\networkshare\folder\CMSReport.htm"


#####################################################################################################################
############################################## Cluster Mailbox Status ###############################################


		
		$Status = Get-ClusteredMailboxServerStatus

			if ($Status.State -eq 'Online') {
		
			$Identity = $status.identity
			$ServerName =  $status.ClusteredMailboxServerName
			$State =  $status.State
			$OperationalMachines = $status.OperationalMachines
			$FailedResources = $Status.FailedResources
			$FailedReplicationHostNames = $Status.FailedReplicationHostNames

			$machines = $operationalMachines[0] + "," + $operationalMachines[1]
			$Machines = $Machines.Replace("<","{")
			$Machines = $Machines.Replace(">","}")

			Add-Content $report "<tr>" 
			Add-Content $report "<td width='10%' bgcolor= 'GainsBoro' align=center>  <B> $Identity</B></td>" 
         		
			Add-Content $report "<td  width='5%' bgcolor= 'Aquamarine' align=center>  <B>$State</B></td>" 
			Add-Content $report "<td  width='20%' bgcolor= 'GainsBoro' align=center>  <B>$Machines</B></td>" 
			Add-Content $report "<td  width='15%' bgcolor= 'GainsBoro' align=center>  <B>$FailedResources</B></td>" 
			Add-Content $report "<td  width='15%' bgcolor= 'GainsBoro' align=center>  <B>$FailedReplicationHostNames</B></td>" 
			
			Add-Content $report "</tr>" 
			
			}


			Else {
			$Identity = $status.identity
			$ServerName =  $status.ClusteredMailboxServerName
			$State =  $status.State
			$OperationalMachines = $status.OperationalMachines
			$FailedResources = $Status.FailedResources
			$FailedReplicationHostNames = $Status.FailedReplicationHostNames
			


			Add-Content $report "<tr>" 
			Add-Content $report "<td width='10%' bgcolor= 'GainsBoro' align=center>  <B> $Identity</B></td>" 
         		
			Add-Content $report "<td width='5%' bgcolor= 'Red' align=center>  <B>$State</B></td>" 
			Add-Content $report "<td width='20%' bgcolor= 'GainsBoro' align=center>  <B>$Machines</B></td>" 
			Add-Content $report "<td width='15%' bgcolor= 'GainsBoro' align=center>  <B>$FailedResources</B></td>" 
			Add-Content $report "<td width='15%' bgcolor= 'GainsBoro' align=center>  <B>$FailedReplicationHostNames</B></td>" 
			
			Add-Content $report "</tr>" 

			}

##############################################################################################################################








<# 
    .SYNOPSIS 
    Removes/Disables HealthMailboxes that are lacking a mailbox database attribute aka corrupt

    Thomas Stensitzki 

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

    Version 1.1, 2016-11-28 

    Please send ideas, comments and suggestions to support@granikos.eu 

    .LINK 
    More information can be found at http://scripts.granikus.eu

    .DESCRIPTION 
    This script removes/disables HealthMailboxes that show an inconsistent error when querying monitoring mailboxes using
		Get-Mailbox -Monitoring
    and receiving a warning like
    
    "WARNING: The object DOMAINNAME/Microsoft Exchange System Objects/Monitoring Mailboxes/�Health_Mailbox_GUID has been corrupted, 
	and it's in an inconsistent state. The following validation errors happened: WARNING: Database is mandatory or UserMailbox.
 
    .NOTES 	
    Requirements 
    - Windows Server 2008 R2 SP1, Windows Server 2012 or Windows Server 2012 R2  
	
	  REMOVE currently does not work as expected
    
    Revision History 
    -------------------------------------------------------------------------------- 
    1.0 Initial community release 
    1.1 Some PowerShell hygiene

    .PARAMETER Remove  
    Remove the HealthMailboxes that have an empty database attribute  

    .PARAMETER Disable
    Disables the HealthMailboxes that have an empty database attribute
    
    .EXAMPLE 
    Remove the HealthMailbox(es) having an empty database attribute
    .\Fix-HealthMailboxes.ps1 -Remove   
#>  

Param(
    [parameter(Mandatory=$true,HelpMessage='Remove HealthMailboxes in a corrupted state',ParameterSetName="R")]
    [switch]$Remove,
    [parameter(Mandatory=$true,HelpMessage='Disable HealthMailboxes in a corrupted state',ParameterSetName="D")]
    [switch]$Disable
)

function script:CheckHealthMailboxes()
{
    $healthMailboxes = Get-Mailbox -Monitoring -WarningAction SilentlyContinue

    $count = $healthMailboxes.Count

    Write-Output "$count HealthMailbox(es) found" 

    $i=0
    
    foreach($mailbox in $healthMailboxes) {
        try {
            Write-Output "Checking:"$mailbox.UserPrincipalName
            
            If(($mailbox.database -eq '') -or ($mailbox.database -eq $null)) {
                $upn = $mailbox.UserPrincipalName

                Write-Warning "Database attribute check failed: $upn" 
                if($i -lt 50) {
                    try {
                        Write-Output "Disabling mailbox $upn"
                        $i++
						
                        if($Remove) {
                            Remove-Mailbox $upn -Confirm:$false -Permanent:$true -ErrorAction Stop
                        }
                        if($Disable) {
                            Disable-Mailbox $upn -Confirm:$false -ErrorAction Stop
                        }
                    }
                    catch {
                        Write-Error "Error deleting mailbox! Please check, if you have sufficient permission to delete the account!"
                    }
                }
            }
        }
        catch {
            $upn = $mailbox.UserPrincipalName
            Write-Warning "Mailbox with warning: $upn"
        }
    }
    Write-Output "$i HealthMailbox(es) deleted"
}

## MAIN ---------------------------------------
Set-ADServerSettings -ViewEntireForest $true

CheckHealthMailboxes

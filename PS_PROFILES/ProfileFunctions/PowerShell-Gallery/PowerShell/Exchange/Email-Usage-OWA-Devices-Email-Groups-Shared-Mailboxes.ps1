<#
CopyLeft SVET :)
.Synopsis
   Check Email Usage, Mailbox Statistics, OWA Devices, Email Groups Members and Shared Mailboxes Permissions
.DESCRIPTION
   Check Email Usage, Mailbox Statistics, OWA Devices, Email Groups Members and Shared Mailboxes Permissions
.EXAMPLE
#>
Function Get-UserMailboxSize
{
    #Check user email usage
    param(
    [CmdletBinding()]
    [Parameter(Mandatory=$true)]
    [String]$emailAddress    
)
    Get-MailboxStatistics $emailAddress | Select-Object -Property DisplayName, TotalitemSize
}
#You might have an error called BlockAFToExternalUser
#
#Get the currently configured a Remote Domain
Get-RemoteDomain
#
#See the current configuration for the remote domain
#
Get-RemoteDomain Default | Format-List
$defaultdomain = Get-RemoteDomain Default | Select-Object -Property DomainName | Format-Table -HideTableHeaders
#
#Look for the settings for AllowedOOFType, AutoReplyEnabled and AutoForwardEnabled
#
Get-RemoteDomain Default | Format-List AllowedOOFType, AutoReplyEnabled, AutoForwardEnabled
#
#Change the setting
#
##Enable automatic replies
#
Set-RemoteDomain * -AutoReplyEnabled $true
#
##Enable automatic forwards
#
Set-RemoteDomain * -AutoForwardEnabled $true
#
##Enable OOF for Outlook
#
Set-RemoteDomain * -AllowedOOFType $ExternalLegacy
#
##To change all at once
#; 
Set-RemoteDomain * -AutoReplyEnabled $true -Verbose; Set-RemoteDomain * -AutoForwardEnabled $true -Verbose; Set-RemoteDomain * -AllowedOOFType $ExternalLegacy -Verbose
#
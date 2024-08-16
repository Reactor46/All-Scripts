<#
 Exchange 2010 Firewall Request Generation Script
 Author:	Zachary Loeber
 Date:		07/29/2011 
 Version:	1.4
 Description:
	Use this script to automatically generate a csv file for required network 
	communication between servers in an exchange environment. This can be 
	particularly nice to have in large organizations with multiple sites
	or a heavily segmented network.
 Requirements:
	The following files need to be present where you save the script:
	FirewallRules.csv"
	ExchangeEnvironment.csv"
 Example:
 	./generatefirewallrules.ps1 -Envfile "./ExchangeEnvironment.csv" 
	    -RuleFile "./FirewallRules.csv" -OutputFile "./FirewallOutput.csv"
 Notes:
	Some roles have been made generic. So for a third party anti-spam vendor you 
	would use the role of "Internet". For a server that is hosting both the 
	Hub-Transport and the Client-Access roles you will need to list the server twice
	(once for each role).

	To cut out some of the unneeded rules I added a few options to skip rule
	generation for systems on the same site as well.

	I welcome recommendations or corrections zloeber (at) gmail (dot) com
 Change Log:
  	1.4 - Fixed some logic around Client-Network processing to generate just rules to the same site for hub-transport/
		  Client-Access roles and to bypass $SkipSameSite settings. 
		- Updated the FirewallRules.csv to be more detailed for setting static ports for cross-site dags
		  (This is actually a really convenient reference in its own right)
		- Added a region column to the exchange environment csv file for processing
 	1.3 - Added logic for client-network rules to only process them if they are in the same site as the Role
		  In our input exchange environment csv file if you want 2 sites to generate rules that allow them
		  to reach two other sites instead of just their own you will need to put the network in twice, once
		  for each site like so:
			Client-Network,10.203.2.0/24,End User Network - Site1,Site1
			Client-Network,10.203.2.0/24,End User Network - Site1,Site2
 	1.2	- Fixed some DAG/database related exceptions
	1.1 - Made entire script callable as a function
	1.0 - Initial release
#>

param([parameter(Mandatory=$true,ValueFromPipeline=$false,HelpMessage="Environment to process")][string]$EnvFile, 
	[parameter(Mandatory=$true,ValueFromPipeline=$false,HelpMessage="Rules to process")][string]$RuleFile,
	[parameter(Mandatory=$true,ValueFromPipeline=$false,HelpMessage="Output CSV File")][string]$OutputFile,
	[parameter(Mandatory=$false,ValueFromPipeline=$false,HelpMessage="Process only cross-site dc rules")][switch]$LocalDCsOnly = $true,	
	[parameter(Mandatory=$false,ValueFromPipeline=$false,HelpMessage="Skip all DAG related rules")][switch]$SkipLocalDAGRuleGeneration = $false,	
	[parameter(Mandatory=$false,ValueFromPipeline=$false,HelpMessage="Skip rules where servers are in the same site")][switch]$SkipSameSite = $true)
	
$Rules = @(Import-Csv $RuleFile)
$ApplicableRules = @()
$FirewallRequests = @()
$FirewallRequest = New-Object Object
$ExchangeEnv = @(Import-Csv $EnvFile)
$ExchangeEnvRoles = @($ExchangeEnv | Select Role -Unique)
$SameSite = $false
$ProcessRule = $false

# First lets get the collection of rules which only applies to the roles in the environment
Foreach ($Rule in $Rules) {
	if ($ExchangeEnvRoles -match [string]$Rule.Destination_Role){
		$ApplicableRules = $ApplicableRules + $Rule	
	}
}

# Loop through each server
Foreach ($ExchangeServer in $ExchangeEnv) {
	# Loop through all servers again to find rules which match against the current server
	Foreach ($ExchangeOtherServer in $ExchangeEnv) {
		# Do nothing if the server is itself essentially
		If (($ExchangeServer.Role + $ExchangeServer.ServerName + $ExchangeServer.IP) -ne  ($ExchangeOtherServer.Role + $ExchangeOtherServer.ServerName + $ExchangeOtherServer.IP)) {
			# Find out if the two servers are in the same site or not
			If ($ExchangeServer.Site -eq $ExchangeOtherServer.Site) {
				$SameSite = $true
			}
			Else {
				$SameSite = $false
			}
			
			# Loop through all the applicable rules
			Foreach ($Rule in $ApplicableRules) {
				# If we are not skiping rules to/from the same site and the rules are not proxy* or internet then continue generating the rule
				If  (-not (($SkipSameSite -and $SameSite) -and `
				(-not (($ExchangeServer.Role -like "Proxy*" -or $ExchangeOtherServer.Role -like "Proxy*") `
				-or ($Rule.Source_Role -eq "Internet") `
				-or ($Rule.Source_Role -eq "Client-Network"))))) {
					# Source and destination roles match so continue processing
					If (($ExchangeServer.Role -eq $Rule.Source_Role) -and ($ExchangeOtherServer.Role -eq $Rule.Destination_Role)) {
						$ProcessRule = $true
						# Only process proxy rules to servers in the same site as the proxy
						If (($Rule.Source_Role -like "Proxy*") -and (-not $SameSite)) {
							$ProcessRule = $false
						}
						# Same thing with edge servers, should be in the same site
						If (($Rule.Source_Role -like "Edge*") -or ($Rule.Destination_Role -like "Edge*") -and (-not $SameSite)) {
							$ProcessRule = $false
						}
						If (($Rule.Source_Role -eq "Client-Access") -and ($Rule.Destination_Role -like "Database*") -and (-not $SameSite)) {
							$ProcessRule = $false
						}
						# Again, same for internet, must be in same "site" as destination server
						If (($Rule.Source_Role -eq "Internet") -and (-not $SameSite)) {
							$ProcessRule = $false
						}
						
						# Do some verification before spitting out rules from the internet to a Hub-Transport server
						If (($Rule.Source_Role -eq "Internet") -and ($Rule.Destination_Role -eq "Hub-Transport") -and ($SameSite)) {
							# I honestly don't know how I came up with this monstrosity but it essentially looks to see if there is an edge
							#  server in the current site or not, if there is then skip Internet to Hub-Transport rules (this mess deserves some revisiting)
							If (($ExchangeEnv -match $ExchangeServer.Site | Where {($_.Role -eq "Edge-Transport")}| `
							select Role -Unique).Role -eq "Edge-Transport") {
								$ProcessRule = $false
							}						
						}
						# Like above, but instead for internet to Client-Access servers
						If (($Rule.Source_Role -eq "Internet") -and ($Rule.Destination_Role -eq "Client-Access") -and ($SameSite)) {
							# Same crazy mess as the prior check, you figure it out if you like.
							If (($ExchangeEnv -match $ExchangeServer.Site | Where {($_.Role -like "Proxy-External")}| `
							select Role -Unique).Role -contains "Proxy-External") {
								$ProcessRule = $false
							}						
						}
						# Skip rule generation if we are referencing a DC in a different site than the server and the 
						#	option $localdcsonly is set to true
						If (($Rule.Destination_Role -eq "DC") -and ($LocalDCsOnly) -and (-not $SameSite)) {
							$ProcessRule = $false
						}
						
						#Skip rule generation if database to database and not in the same DAG
						If (($ExchangeServer.Role -like "Database*") -and ($ExchangeOtherServer.Role -like "Database*") -and `
						($ExchangeServer.DAG_Name -ne $ExchangeOtherServer.DAG_Name)) {
							$ProcessRule = $false
						}
						
						#Skip rule generation if database to database and in the same DAG/Site but we choose not to generate rules
						If (($ExchangeServer.Role -like "Database*") -and ($ExchangeOtherServer.Role -like "Database*") -and `
						($ExchangeServer.DAG_Name -eq $ExchangeOtherServer.DAG_Name) -and ($SameSite) -and ($SkipLocalDAGRuleGeneration)) {
							$ProcessRule = $false							
						}
						
						#Skip rule generation if database to non-database and not same
						If ((($ExchangeServer.Role -like "Database*") -or ($ExchangeOtherServer.Role -like "Database*")) -and `
						(($ExchangeServer.Role -notlike "Database*") -or ($ExchangeOtherServer.Role -notlike "Database*")) -and (-not $SameSite)) {
							$ProcessRule = $false							
						}

						#Skip rule generation if Client-Network talking to Client-Access and not in same site
						If (($Rule.Source_Role -eq "Client-Network") -and ($Rule.Destination_Role -eq "Client-Access") `
						-and (-not $SameSite)) {
							$ProcessRule = $false							
						}						
						#Skip rule generation if Client-Network talking to Hub-Transport and not in same site
						If (($Rule.Source_Role -eq "Client-Network") -and ($Rule.Destination_Role -eq "Hub-Transport") `
						-and (-not $SameSite)) {
							$ProcessRule = $false							
						}
						If ($ProcessRule) {
							$FirewallRequest | Add-Member NoteProperty "Source Server" $ExchangeServer.ServerName;
							$FirewallRequest | Add-Member NoteProperty "Source Server IP" $ExchangeServer.IP;
							$FirewallRequest | Add-Member NoteProperty "Source Server Site" $ExchangeServer.Site;
							$FirewallRequest | Add-Member NoteProperty "Destination Server" $ExchangeOtherServer.ServerName;
							$FirewallRequest | Add-Member NoteProperty "Destination Server IP" $ExchangeOtherServer.IP;
							$FirewallRequest | Add-Member NoteProperty "Destination Server Site" $ExchangeOtherServer.Site;
							$FirewallRequest | Add-Member NoteProperty "Port" $Rule.Port;
							$FirewallRequest | Add-Member NoteProperty "Protocol" $Rule.Protocol;
							$FirewallRequest | Add-Member NoteProperty "Description" $Rule.Description;
							$FirewallRequest | Add-Member NoteProperty "Default Authentication" $Rule.DefaultAuthentication;
							$FirewallRequest | Add-Member NoteProperty "Supported Authentication" $Rule.SupportedAuthentication;
							$FirewallRequest | Add-Member NoteProperty "Encryption Supported" $Rule.EncryptionSupported;
							$FirewallRequest | Add-Member NoteProperty "Encrypted by Default" $Rule.EncryptedbyDefault;
							$FirewallRequest | Add-Member NoteProperty "Source Role" $ExchangeServer.Role;
							$FirewallRequest | Add-Member NoteProperty "Destination Role" $ExchangeOtherServer.Role;
							$FirewallRequest | Add-Member NoteProperty "Source Region" $ExchangeServer.Region;
							$FirewallRequest | Add-Member NoteProperty "Destination Region" $ExchangeOtherServer.Region;
							$FirewallRequest | Add-Member NoteProperty "Notes" $Rule.Notes;			
							$FirewallRequests += $FirewallRequest
							$FirewallRequest = New-Object Object
						}
					}
				}
			}
		}
	}	
}


$FirewallRequests | select * | Export-Csv -NoTypeInformation $OutputFile
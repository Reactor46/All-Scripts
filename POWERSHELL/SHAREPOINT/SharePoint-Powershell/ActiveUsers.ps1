# Execute export command
# ./ActiveUsers.ps1 -configurationFile "C:\inetpub\wwwroot\wss\VirtualDirectories\33843\web.config" -membershipProviderName "FBADrugRepMembershipProvider" -outputFile Results.txt
#
param
( 
	# Configuration file
	[alias("configFile")]
   	[string] $configurationFile = $(throw "The parameter 'configurationFile' is required."),
	
	# Membershp provider
	[alias("membershpiProvider")]
    [string] $membershipProviderName = $(throw "The parameter 'membershipProviderName' is required."),

	# Output file
	[alias("output")]
   	[string] $outputFile = $(throw "The parameter 'outputFile' is required."),

	# Output file
   	[switch] $overwrite
)

function Resolve-Path-With-Force($filename)
{  
	$filename = Resolve-Path $filename -ErrorAction SilentlyContinue -ErrorVariable _frperror  
	
	if (!$filename)  
	{    
		return $_frperror[0].TargetObject  
	}  
	return $filename
}

# Resolve paths
$configurationFile = Resolve-Path-With-Force $configurationFile
$outputFile = Resolve-Path-With-Force $outputFile

# Does the configuration file exist
if(![System.IO.File]::Exists($configurationFile))
{
	throw "The configuration file does not exist."
}

# Clear the output file
if([System.IO.File]::Exists($outputFile))
{
	if($overwrite)
	{
		Remove-Item $outputFile
	}
	else
	{
		throw "the output file '$outputFile' already exists. Use the [overwrite] flag to replace the existing file."
	}
}

# Setup application configuration
[System.AppDomain]::CurrentDomain.SetData("APP_CONFIG_FILE", $configurationFile) 
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Configuration")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Web") 

# Begin
Write-Host("`n------ Export started ------ ")

# Get the default provider
Write-Host "Getting the '$membershipProviderName' membership provider."
$membershipProvider = [System.Web.Security.Membership]::Providers[$membershipProviderName]

# Clear the log file
if($membershipProvider -eq $null)
{
	throw "The role provider '$membershipProviderName' could not be found."
}

# Set the page index
$pageIndex = 0
$pageSize = 100
$numberOfUsersTotal = 0

# Add output header
Add-Content $outputFile "Active Accounts"

	[System.Collections.ArrayList]$caps = @("betty.anderson@boehringer-ingelheim.com", "michelle.billberry@boehringer-ingelheim.com","perry.brown@boehringer-ingelheim.com",
	"coalson_bryana@allergan.com","michele.byers@novartis.com","amy.carter@boehringer-ingelheim.com","lisaclissold@yahoo.com","tracy.coker@bayer.com","robert.cole@gilead.com",
	"econner@celgene.com","skylerdalton@hotmail.com","anliau1010@gmail.com","claudia.friedman@bayer.com","tracy.e.goswick@gsk.com","christine.hajovsky@boehringer-ingelheim.com",
	"amy.hauff@boehringer-ingelheim.com","tiffany.hedman@novartis.com","hmcesham01@gmail.com","kelli.waddle@gsk.com","joseph@thejamsession.net","melinda.neese@boehringer-ingelheim.com",
	"april.obrien@bayer.com","debra.patrick1@astrazeneca.com","chadoprince@sbcglobal.net","lorri.d.steen@gsk.com","ctascher@its.jnj.com","hemang.thakkar@abbvie.com","blakeley.r.thomas@gsk.com",
	"kristin.thomas@bayer.com","kimberly.williams@boehringer-ingelheim.com")

	$deleteList = New-Object System.Collections.ArrayList

do
{
    # Get users in page
    $users = $membershipProvider.GetAllUsers($pageIndex, $pageSize, [ref] $numberOfUsersTotal)

    # Calculate the beginning user index
    $userIndex = ($pageIndex * $pageSize);

	# Setup the number of users
	if($users -ne $null)
	{
		$numberOfUsersInPage = $users.Count
	}
	else
	{
		$numberOfUsersInPage = 0
	}
		# Process users
		foreach ($user in $users)
		{
			if($caps -contains $user.UserName)
			{
					# Get user properties
					$username = $user.UserName
					$email = $user.Email

					# add the user to the output file
					Add-Content $outputFile "$username,$email"
			}
			else
			{
					$username = $user.UserName
					$deleteList += $user.UserName
					#Add-Content $outputFile "$username"
			}
		}
		
    # Goto next page
    $pageIndex++
}
while ($numberOfUsersInPage -eq $pageSize)

for ($i=0; $i -lt $deleteList.Count;$i++)
		{
			$membershipProvider.DeleteUser($deleteList[$i],$true)
			# Write processing statement
					write-host([System.String]::Format(
						"Deleting : {0}",
						$deleteList[$i]));
		}
 
# Finish message
Write-Host("==========  Finished export: $numberOfUsersTotal users ========== `n")
#Read more: https://www.sharepointdiary.com/2016/04/create-enterprise-search-center-using-powershell.html#ixzz8BFBsal7K
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 
#Define Variables for Web Application Creation
$WebAppName = "Crescent Search Center"
$HostHeader = "search.crescent.com"
$WebAppURL = "https://" + $HostHeader
$WebAppPort = "80"
$ContentDBName = "Crescent_Search_Content"
$AppPoolName = "Crescent Search Web Application App Pool"
$AppPoolAccount = "Crescent\SP16_Pool"
$FarmAdminAccount = "Crescent\SP16_Farm"
 
#Authentication Provider
$AuthProvider = New-SPAuthenticationProvider
 
#Check if Managed account is registered already
Write-Host -ForegroundColor Yellow "Checking if the Managed Accounts already exists"
$AppPoolAccount = Get-SPManagedAccount -Identity $AppPoolAccount -ErrorAction SilentlyContinue
if ($AppPoolAccount -eq $null)
{
    Write-Host "Please Enter the password for the Service Account..."
    $AppPoolCredentials = Get-Credential $AppPoolAccount
    $AppPoolAccount = New-SPManagedAccount -Credential $AppPoolCredentials
}
 
#Create new Web Application
New-SPWebApplication -name $WebAppName -port $WebAppPort -hostheader $HostHeader -URL $WebAppURL -ApplicationPool $AppPoolName -ApplicationPoolAccount (Get-SPManagedAccount $AppPoolAccount) -AuthenticationMethod NTLM -AuthenticationProvider $AuthProvider -DatabaseName $ContentDBName
 
#Create Enterprise Search Center site collection
New-SPSite -Name $WebAppName -Url $WebAppURL -Template "SRCHCEN#0" -OwnerAlias $FarmAdminAccount -ContentDatabase $ContentDBName

<#
Next step: Configure search center permissions
Once the search center site is created, you must grant permission to all users in the organization to access the search center site. 

Go to : Site Settings >> Site Permissions >> Enterprise Search Visitors Group >> Add “Everyone”
#>
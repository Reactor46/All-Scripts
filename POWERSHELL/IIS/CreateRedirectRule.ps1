##############################
#.SYNOPSIS
# Creates a HTTP to HTTPS URL redirect rule in the given site's web.config
#
#.DESCRIPTION
# Creates a HTTP to HTTPS URL redirect rule in the given site's web.config
#
#.PARAMETER SiteBinding
# Value of the target sites IIS binding.
#
# REQUIRED
#
# ex: coffeehouse.thecodattic.com
#
#.PARAMETER RuleName
# Name to be assigned to the redirect rule that is created.
#
# REQUIRED
#
# ex: 'HTTP to SSL Redirect'
#
#.EXAMPLE
#
# > .\CreateRedirectRule.ps1 -SiteBinding "coffeehouse.thecodeattic.com" -RuleName "Redirect to SSL"
#
#.EXAMPLE
#
# > .\CreateRedirectRule.ps1 "coffeehouse.thecodeattic.com" "Redirect to SSL"
#
#.NOTES
# Creates a HTTP to HTTPS URL redirect rule in the given site's web.config
##############################
param( 
    [parameter(Position=0, Mandatory=$true, HelpMessage="Value of the target sites IIS binding.")]
    [string]$SiteBinding,
    [parameter(Position=1, Mandatory=$true, HelpMessage="Name to be assigned to the redirect rule that is created.")]
    [string]$RuleName
)

Add-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST/$SiteBinding" -filter "/system.webServer/rewrite/rules" -name "." -value @{name=$RuleName;patternSyntax='Wildcard';stopProcessing='True';}
Add-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST/$SiteBinding" -filter "/system.webServer/rewrite/rules/rule[@name='$RuleName']/conditions" -name "." -value @{logicalGrouping="MatchAny";input="{HTTPS}";pattern="off";}
Set-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST/$SiteBinding" -filter "/system.webServer/rewrite/rules/rule[@name='$RuleName']/match" -name "url" -value "*"
Set-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST/$SiteBinding" -filter "/system.webServer/rewrite/rules/rule[@name='$RuleName']/action" -name "type" -value "Redirect"
Set-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST/$SiteBinding" -filter "/system.webServer/rewrite/rules/rule[@name='$RuleName']/action" -name "url" -value "https://{HTTP_HOST}{REQUEST_URI}"
Set-WebConfigurationProperty -pspath "MACHINE/WEBROOT/APPHOST/$SiteBinding" -filter "/system.webServer/rewrite/rules/rule[@name='$RuleName']/action" -name "redirectType" -value "Found"

Write-Host "$RuleName has been created for $SiteBinding" -BackgroundColor DarkGreen -ForegroundColor Gray
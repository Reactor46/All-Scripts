# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

# Synopsis: Create Eas RegisterNotification extension virtual directory on CAFE server
#           

param(
  [bool]$FrontEnd = $(throw "Value for parameter FrontEnd not provided.  Acceptable values are true or false.")
)

############################
# Common function for error logging
############################
function Log-Error ( [string]$errorMsg, [bool]$Stop=$true )
{
    Write-ExchangeSetupLog -Error $errorMsg;
    
    if ($Stop)
    {
        throw "$errorMsg" 
    }
}

############################
# Common function for creating a site
############################
function Setup-Site ( [string]$siteName, [string]$relativePhysicalPath)
{
    Write-ExchangeSetuplog -info "Trying to setup EasNotificationExtension on $siteName";
    
    $easAppPoolName = "MSExchangeSyncAppPool";

    $libPath = "$env:SystemRoot\system32\inetsrv\microsoft.web.administration.dll";
    add-type -Path $libPath;
    $serverMgr = new-object Microsoft.Web.Administration.ServerManager;

    $easAppPool = $serverMgr.ApplicationPools[$easAppPoolName];
    if ($easAppPool -eq $null)
    {
        Log-Error "$easAppPoolName does not exist";
    }
    
    $iisSite = $serverMgr.Sites[$siteName];
    if ($iisSite -eq $null)
    {
        Log-Error "$siteName does not exist";
    }
    
    $easApp = $iisSite.Applications["/Microsoft-Server-ActiveSync"];
    if ($easApp -eq $null)
    {
        Log-Error "/Microsoft-Server-Activesync virtual directory does not exist";
    }
    
    $sapiApp = $iisSite.Applications["/Sapi"];
    if ($sapiApp -ne $null)
    {
        Write-ExchangeSetupLog -Info "/Sapi virtual directory already exists.  Deleting it and recreating";
        $sapiApp = $iisSite.Applications.Remove($sapiApp);
    }
            
    $physicalPath = [System.IO.Path]::Combine($RoleInstallPath, $relativePhysicalPath);

    Write-ExchangeSetupLog -Info "Creating /Sapi virtual directory";

    $sapiApp = $iisSite.Applications.Add("/Sapi", $physicalPath);
    $sapiApp.ApplicationPoolName = $easAppPool.Name ; 

    Write-ExchangeSetuplog -Info "Updating IIS ApplicationHostConfig";

    $appConfig = $ServerMgr.psbase.GetApplicationHostConfiguration();

    Write-ExchangeSetuplog -info "Configuring authentication methods";

    $appConfig.GetSection("system.webServer/security/authentication/anonymousAuthentication","$siteName/Sapi").SetAttributeValue("enabled", $false);
    $appConfig.GetSection("system.webServer/security/authentication/basicAuthentication","$siteName/Sapi").SetAttributeValue("enabled", $false);
    $appConfig.GetSection("system.webServer/security/authentication/windowsAuthentication","$siteName/Sapi").SetAttributeValue("enabled", $false);
    $appConfig.GetSection("system.webServer/security/authentication/digestAuthentication","$siteName/Sapi").SetAttributeValue("enabled", $false);
    $appConfig.GetSection("system.webServer/security/access","$siteName/Sapi").SetAttributeValue("sslFlags", 264);
    
    Write-ExchangeSetuplog -Info "Committing changes";
    $ServerMgr.CommitChanges();
}


############################
# Main execution starts here
############################
Write-ExchangeSetupLog -Info "Install-EasNotificationExtensionVirtualDirectory beginning with FrontEnd value $FrontEnd";

if ($FrontEnd)
{
    Setup-Site "Default Web Site" "FrontEnd\HttpProxy\sync";
}
else
{
    Setup-Site "Exchange Back End" "ClientAccess\sync";
}

Write-ExchangeSetuplog -Info "Install-EasNotificationExtensionVirtualDirectory finished.";


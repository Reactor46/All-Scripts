<#
    This runbook demonstrates how you could monitor websites in your Azure subscription and alert the 
    on call engineer of any issues via text message, as demonstrated in the Azure Automation announcement.
    It is intended for demonstration purposes only and is not meant for production use (see disclaimer at bottom of
    this comment block).
    
    In addition to this runbook, you will need an OrgID credential with access to your subscription
    (http://azure.microsoft.com/blog/2014/08/27/azure-automation-authenticating-to-azure-using-azure-active-directory/) 
    stored in a credential asset, a Twilio account, and the Twilio PowerShell module.

    In Azure Automation:
     - Import the Twilio PowerShell module (https://gallery.technet.microsoft.com/scriptcenter/Twilio-PowerShell-Module-8a8bfef8)
     - Import this runbook
     - Create a connection asset for Twilio, with your Twillio AccountSid and authtoken. These can be obtained on the Twilio website
     - Update this runbook to contain the name of your Azure OrgID credential asset, subscription, and Twilio connection asset.
     - Start this runbook, providing a value for the OnCallPhone parameter.

    To cause the runbook to send a text message, stop one of your Azure websites.

    WARNING: This runbook loops continually and never stops. To avoid using up a lot of job execution time, 
    DO NOT LEAVE IT RUNNING! If you want to use this runbook in production scenarios, modify it to only perform a check once,
    and execute it on a schedule. Or consider modifying it to fit with the monitoring solution described here:
    http://azure.microsoft.com/blog/2014/11/17/monitoring-azure-services-and-external-systems-with-azure-automation/
#>
workflow Monitor-AzureWebsite
{
    param(
        [string] $OnCallPhone
    )

    $TwilioCon = Get-AutomationConnection -Name 'joeTwilio' # Replace with the name of your Twilio connection asset
    $AzureOrgIdCredential = Get-AutomationPSCredential -Name 'joeAzureCred' # Replace with the name of your Azure OrgID cred asset
    $AzureSubscriptionName = "Windows Azure MSDN - Visual Studio Ultimate" # Replace with the name of your Azure subscription

    $Null = Add-AzureAccount -Credential $AzureOrgIdCredential  
    $Null = Select-AzureSubscription -SubscriptionName $AzureSubscriptionName

    $LastWebsiteStatus = @{}

    "Monitoring Azure Websites..."

    while($True) {
        $Websites = Get-AzureWebsite
        
        $Websites | ForEach-Object {
            if($_.State -ne "Running" -and $LastWebsiteStatus[$_.Name] -ne $_.State) {
                "Website " + $_.Name + " is not responding! Alerting on call engineer."
                   
                $FromPhone = Get-TwilioPhoneNumbers -Connection $TwilioCon                   
                   
                # send text message
                $Output = Send-TwilioSMS `
                    -Connection $TwilioCon `
                    -From $FromPhone.PhoneNumber `
                    -Message ("Website " + $_.Name + " is not responding. Go fix it!") `
                    -To $OnCallPhone 
            }
               
            $LastWebsiteStatus[$_.Name] = $_.State
        }
        
        "Sleeping..."
        Start-Sleep -Seconds 5
    }
    
}
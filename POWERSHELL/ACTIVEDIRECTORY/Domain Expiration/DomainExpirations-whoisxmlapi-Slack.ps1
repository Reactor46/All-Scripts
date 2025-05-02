# Static variables
$uri = Get-AutomationVariable -Name 'DomainExpirationSlackHook'
$countdown = -60 #maximum number of days before domain expires to be notified by
$whoisxml = New-Object System.Xml.XmlDocument
#whoisxmlapi.com account
$apiuser = Get-AutomationVariable -Name 'WhoIsApiUser'
$apipass = Get-AutomationVariable -Name 'WhoIsApiPass'
 
#Defines credentials for authenticating to clients
$office365creds = Get-AutomationPSCredential -Name 'Office365ScriptingCreds'
# End automation.

# Used to run once.
#$office365creds = Get-Credential
# End run once.

Write-Output "Connecting to session..."
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $office365creds -Authentication Basic -AllowRedirection
Import-PSSession $session -ErrorAction 'silentlycontinue'

if ($?) {
    Write-Output "Connected to session!`n`n"
    Write-Output "Connecting to MsolService..."
    Connect-MsolService -Credential $office365creds
   

    if ($?) {
        Write-Output "Connected to MsolService!`n`n"
        Write-Output "Getting tenants..."
        # Getting a list of tenant IDs (clients) used to connect to their environment
        $tenants = Get-MsolPartnerContract
        Write-Output "Tenants retrieved!`n`n"

        $loop = 0 #used to rotate colors per client

        # Running command against all tenants
        ForEach ($tenant in $tenants) {   
            $slackmessage = @{}
            $attachment = @{}
            $domain_field = @{}
            $date_field = @{}

            # Cycles through colors so every client doesn't look the same
            if ($loop % 2 -eq 0 ) {
                $color = "#0066ff"
            }
            else {
                $color = "#36a64f"
            }

            $domains_list = "" # Will be a list of domains to add to a field in Slack
            $dates_list = "" # Will be a list of dates to add to a field in Slack

            $TenantID = $tenant.TenantID
            $TenantName = $tenant.Name

            Write-Output "`n`nProcessing $TenantName...`n"

            $domains = Get-MsolDomain -TenantID $TenantID | Where { $_.Name -NotLike "*onmicrosoft.com" -and $_.Name -NotLike "*microsoftonline.com" }
            
            ForEach ($domain in $domains) {
                $domainname = $domain.Name
                
                Write-Output "Processing $domainname..."

                $whoisxml.Load("http://www.whoisxmlapi.com/whoisserver/WhoisService?domainName=$domainname&username=$apiuser&password=$apipass")
                
                if ($domainname -like "*.org" -and $domainname) {
                    Write-Output "Domain is a .org"
                    
                    $date = $whoisxml.WhoisRecord.registryData.ExpiresDate
                }
                else {
                    $date = $whoisxml.WhoisRecord.ExpiresDate
                }
                
                # Converts $date variable to a string. For some reason, the response from "Select-String"
                # can't be worked on until that happens.
                $date = [DateTime]$date
                
                # NOTICE would appear if the domain didn't have any valid information.
                if ($date -ne "" -and $date) {
                    Write-Output "$domainname will expire on $date"

                    if ((Get-Date) -gt $date.AddDays($countdown) ) {
                        $date = $date.ToString("MM-dd-yyyy")

                        $domains_list += "$domainname`n"
                        $dates_list += "$date`n"
                        
                        Write-Output "$domainname has been queue to send a notification since it will be expiring, soon." 
                    }                
                }
                else {
                    $date = "N/A"

                    $domains_list += "$domainname`n"
                    $dates_list += "$date`n"

                    Write-Output "Date not found for $domainname."
                }
            }
            
            # Validates that domains have expired or aren't responding, adds them to the main body.
            if ($domains_list -ne "") {
                $loop++ # Keeps track of how many clients are actually found.

                <#
                    # This is where the $TenantName header will be added in Slack.
                #>

                $attachment.Add("fallback", "$TenantName has domains expiring.")
                $attachment.Add("color", "$color")
                $attachment.Add("title", "$TenantName")

                <#
                    # This is where the domain and it's expiration should be sent to Slack.
                    # This can be formatted the same even though the dates are invalid.
                #>

                $domain_field.Add("title", "Domain")
                $domain_field.Add("value", "$domains_list")
                $domain_field.Add("short", "true")

                $date_field.Add("title", "Date")
                $date_field.Add("value", "$dates_list")
                $date_field.Add("short", "true")

                $attachment.Add("fields", ($domain_field, $date_field))
                $slackmessage.Add("attachments", @($attachment))

                $slackmessage = $slackmessage | ConvertTo-JSON -Depth 4

                Invoke-WebRequest `
                    -uri "$uri" `
                    -Method POST `
                    -Body $slackmessage

                Write-Host "`Slack message has been sent for $TenantName."
            }
        }

        Get-PSSession | Remove-PSSession
    }
    else {
        Write-Error "Unable to connect to MSOnline module."
    }
}
else {
    Write-Error "Unable to connect to session."
}


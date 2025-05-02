# Static variables
$uri    = "CUSTOM INTEGRATION URL FOR SLACK"
$countdown = -9000 #maximum number of days before domain expires to be notified by


# Needed if running in a scheduled task as system. Otherwise, the root will start in C:\Windows\System32.
#$whoislocation = "C:\WHOIS.EXE LOCATION\"
#cd $whoislocation

# Used to automate.
#$live_user = "O365 USER"
#$live_pass = "O365 PASS"
#$live_pass = ConvertTo-SecureString $live_pass -AsPlainText -Force 
 
#Defines credentials for authenticating to clients
#$office365creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $live_user, $live_pass
# End automation.

# Used to run once.
$office365creds = Get-Credential
# End run once.

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $office365creds -Authentication Basic -AllowRedirection
Import-PSSession $Session -ErrorAction 'silentlycontinue'
Connect-MsolService -Credential $office365creds

# Getting a list of tenant IDs (clients) used to connect to their environment
$tenants = Get-MsolPartnerContract

$loop = 0 #used to rotate colors per client

# Running command against all tenants
ForEach ($tenant in $tenants) 
{   
    $slackmessage = @{}
    $attachment = @{}
    $domain_field = @{}
    $date_field = @{}

    # Cycles through colors so every client doesn't look the same
    if($loop % 2 -eq 0 ) {
        $color = "#0066ff"
    }
    else {
        $color = "#36a64f"
    }

    $color = "#0066ff"

    $domains_list = "" # Will be a list of domains to add to a field in Slack
    $dates_list = "" # Will be a list of dates to add to a field in Slack

    $TenantID = $tenant.TenantID
    $TenantName = $tenant.Name

    Write-Host "Processing $TenantName...`n"

    $domains = Get-MsolDomain -TenantID $TenantID | Where { $_.Name -NotLike "*onmicrosoft.com" -and $_.Name -NotLike "*microsoftonline.com" }
    
    ForEach ($domain in $domains)
    {
        $domainname = $domain.Name
        
        Write-Host "Processing $domainname..."

        # WhoIsCL responds with different information depending on if it's a .org or something else. 
        if($domainname -like "*.org" -and $domainname)
        {
            Write-Host "Domain is a .org" -ForegroundColor "Yellow"
            
            # whois.exe gets hung if being opened in rapid succession. Seems to only happen with .org.
            Write-Host "`nWaiting..."
            Start-Sleep -s 5
            
            $date = .\WhoIs.exe -v $domainname | Select-String -Pattern "Registry Expiry Date: " -AllMatches
        }
        else
        {
            $date = .\WhoIs.exe -v $domainname | Select-String -Pattern "   Expiration Date: " -AllMatches
        }
        
        # Converts $date variable to a string. For some reason, the response from "Select-String"
        # can't be worked on until that happens.
        $date = [String]$date
        
        # NOTICE would appear if the domain didn't have any valid information.
        if($date -notlike "*NOTICE:*" -and $date -ne "" -and $date)
        {
            Write-Host "$date" -ForegroundColor "Green"
            
            $date = $date.Replace("Registry Expiry Date: ", "") # Applies to .org domains
            $date = $date.Replace("   Expiration Date: ", "") # Applies to all other domains
            
            $date = [DateTime]$date

            if((Get-Date) -gt $date.AddDays($countdown) )
            {
                $date = $date.ToString("MM-dd-yyyy")

                $domains_list += "$domainname`n"
                $dates_list += "$date`n"
                
                Write-Host "$domainname will expire on $date." -ForegroundColor "Yellow" 
            }                
        }
        else 
        {
            $date = "N/A"

            $domains_list += "$domainname`n"
            $dates_list += "$date`n"

            Write-Host "Date not found for $domainname." -ForegroundColor "Red"
        }
    }
    
    # Validates that domains have expired or aren't responding, adds them to the main body.
    if($domains_list -ne "")
    {
        $loop++ # Keeps track of how many clients are actually found.

        <#
            # This is where the $TenantName header will be added in Slack.
        #>

        $attachment.Add("fallback","$TenantName has domains expiring.")
        $attachment.Add("color","$color")
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

        $attachment.Add("fields", ($domain_field,$date_field))
        $slackmessage.Add("attachments", @($attachment))

        $slackmessage = $slackmessage | ConvertTo-JSON -Depth 4

        Invoke-WebRequest `
            -uri "$uri" `
            -Method POST `
            -Body $slackmessage

        Write-Host "`Slack message has been sent for $TenantName."
        
        # whois.exe gets hung if being opened in rapid succession.
        Write-Host "`nWaiting..."
        Start-Sleep -s 5
    }
}

Remove-PSSession -Session $Session
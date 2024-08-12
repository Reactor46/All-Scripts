#Static Variables
$email = "SENDER ADDRESS"
$smtp = "SMTP SERVER"
$smtp_user = "SMTP USER"
$smtp_pass = "SMTP PASS"
$recipient = "RECIPIENT ADDRESS"

$whoisxml = New-Object System.Xml.XmlDocument
#whoisxmlapi.com account
$apiuser = "API USER"
$apipass = "API PASS"

#Office 365 Credentials
$live_user = "O365 USER"
$live_pass = "O365 PASS"

$smtp_pass = ConvertTo-SecureString $smtp_pass -AsPlainText -Force
$live_pass = ConvertTo-SecureString $live_pass -AsPlainText -Force
    
$smtpcred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $smtp_user, $smtp_pass
    
#Defines credentials for authenticating to clients
$office365creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $live_user, $live_pass

# Used to run once. Comment out the above to use this.
#$office365creds = Get-Credential
# End run once.


$countdown = -9000 #maximum number of days before domain expires to be notified by
$subject = "Domain Expiration Report -" + [String]$countdown + " Days"

$DateReached = $false
#end Static Variables

$title = @'
<!DOCTYPE html>
<html>
    <head>
        <title>
'@


#CSS and opening body tag
$css =
@'
</title>
        <style type="text/css" media="screen">
            * {
                font-family: Helvetica;
            }

                h2 {
                color: #0066ff;
            }

                h4 {
                font-style: bold;
            }

                table {
                width: 100%;
            }

                td {
                padding: 10px;
            }

                .table_header table {
                border: 1px solid black;
                border-collapse: collapse;
                width: 100%
            }

                .table_header td {
                padding: 10px;
                border: 1px solid black;
            }

               .table_summary {
                width: 50%;
            }

            </style>
    </head>
    <body>
'@

$domain_header_html = @'
    <div class="table_header">
        <table>
            <tr style="font-weight:bold">
                <td width="75%">Domain Name</td>
                <td width="25%"><center>Expiration</center></td>
            </tr>
        </table> 
    </div><br><table>
'@

$body = $title + $subject + $css

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $office365creds -Authentication Basic -AllowRedirection
Import-PSSession $Session -ErrorAction 'silentlycontinue'
Connect-MsolService -Credential $office365creds

#Getting a list of tenant IDs (clients) used to connect to their environment
$tenants = Get-MsolPartnerContract


#Running command against all tenants
ForEach ($tenant in $tenants) 
{   
    #Responsible for holding per-tenant HTML information.
    $tenant_body = ""
    $tenant_domains_expired = 0
    $date = ""

    $TenantID = $tenant.TenantID
    $TenantName = $tenant.Name

    Write-Host "Processing $TenantName...`n"


    #Creating tenant header
    $tenant_body += "<h2>" + $TenantName + "</h2><br>"
        
    #Adding domain header
    $tenant_body += $domain_header_html
    
    $domains = Get-MsolDomain -TenantID $TenantID | Where `
        {$_.Name -NotLike "*onmicrosoft.com" -and $_.Name -NotLike "*microsoftonline.com"}
    
    ForEach ($domain in $domains)
    {
        $domainname = $domain.Name
        
        Write-Host "Processing $domainname..."
        $whoisxml.Load("http://www.whoisxmlapi.com/whoisserver/WhoisService?domainName=$domainname&username=$apiuser&password=$apipass")

        if($domainname -like "*.org" -and $domainname)
        {
            Write-Host "Domain is a .org" -ForegroundColor "Yellow"
            
            $date = $whoisxml.WhoisRecord.registryData.ExpiresDate
        }
        else
        {
            $date = $whoisxml.WhoisRecord.ExpiresDate
        }
        
        $date = [DateTime]$date
        
        if($date -ne "" -and $date)
        {
            Write-Host "$date" -ForegroundColor "Green"

            if((Get-Date) -gt $date.AddDays($countdown) )
            {
                #Once flagged, ensures email is sent to notify. If not flagged, no email will be sent.
                $DateReached = $true
                
                #Once more than 0, tenant will be added to the list.
                $tenant_domains_expired++
                $date = $date.ToString("MM-dd-yyyy")
                $tenant_body += '<tr><td width="75%">' + $domainname + "</td>" + '<td width="25%"><center>' + $date + "</center></td></tr>" 
                
                Write-Host "$domainname will expire on $date." -ForegroundColor "Yellow" 
            }                
        }
        else 
        {
            $tenant_domains_expired++
            $date = "N/A" 
            $tenant_body += '<tr><td width="75%">' + $domainname + "</td>" + '<td width="25%"><center>' + $date + "</center></td></tr>" 
            
            Write-Host "Date not found." -ForegroundColor "Red"
        }
    }
    
    #Validates that domains have expired or aren't responding, adds them to the main body.
    if($tenant_domains_expired -gt 0)
    {
        $DateReached = $true
        
        Write-Host "`Content has been added to the main body."
        $body += $tenant_body + "</table>"
        $body += "<br><br><br>"
    }
    
    Write-Host "`n`n`n" -ForegroundColor "Red"
}

$html_close = @'
    </body>
    </html>
'@

$body += $html_close

if($DateReached -eq $true)
{
    Write-Host "`nMail is being sent to $recipient..."
    Send-MailMessage -To $recipient -Subject $subject -Body $body -BodyAsHtml -From $email -SmtpServer $smtp -usessl -Credential $smtpcred
}

Remove-PSSession -Session $Session
#This script will check to ensure the message queues on LASEXCAS01 and 02 are below 25 messages
$link='"<a href="\\Contosocorp\share\shared\IT\SupportServices\NOC\SOPs\Individual\LASEXCAS01-LASEXCAS02 QUEUE MONITORING.docx">\\Contosocorp\share\shared\IT\SupportServices\NOC\SOPs\Individual\LASEXCAS01-LASEXCAS02 QUEUE MONITORING.docx</a>"'

$exitcode=1

get-pssession | remove-pssession

$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://lasexcas01.Contoso.corp/powershell" -Authentication Kerberos 

Import-PSSession $session -disableNameChecking


    
     Get-Queue | select status,messagecount | %{if($_.status -like "Suspended Queue" -or $_.messagecount -gt "50"){Send-MailMessage -From "noc@creditone.com" -to "noc@creditone.com" -Subject "Critical Alert - EXCHANGE OUTBOUND QUEUE OVER 50 MESSAGES" -Body "Critical Alert - LASEXCAS01 QUEUE <br></br>Please refer to the LASEXCAS01-LASEXCAS02 QUEUE MONITORING SOP located here:<br></br>$link" -BodyAsHtml -Priority High -SmtpServer "mailgateway.Contoso.corp"}}

     Get-Queue -server lasexcas02 | select status,messagecount | %{if($_.status -like "Suspended Queue" -or $_.messagecount -gt "50"){Send-MailMessage -From "noc@creditone.com" -to "noc@creditone.com" -Subject "Critical Alert - EXCHANGE OUTBOUND QUEUE OVER 50 MESSAGES"  -Body "Critical Alert -  LASEXCAS02 QUEUE <br></br>Please refer to the LASEXCAS01-LASEXCAS02 QUEUE MONITORING SOP located here:<br></br>$link" -BodyAsHtml -Priority High -SmtpServer "mailgateway.Contoso.corp"}}

$exitcode 
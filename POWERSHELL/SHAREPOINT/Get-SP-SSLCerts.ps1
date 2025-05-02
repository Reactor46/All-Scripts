Add-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue 
#The mail address of who will receive the backup exception message 
$from = "from@domain.com" 
#Send email function 
function SendMail($subject, $body, $file) 
{ 
  try 
  { 
    #Getting SMTP server name and Outbound mail sender address 
    $caWebApp = (Get-SPWebApplication -IncludeCentralAdministration) | ? { $_.IsAdministrationWebApplication -eq $true } 
    $smtpServer = $caWebApp.OutboundMailServiceInstance.Server.Address 
    $smtp = new-object Net.Mail.SmtpClient($smtpServer) 
    #Creating a Mail object 
    $message = New-Object System.Net.Mail.MailMessage 
    $att = New-Object System.Net.Mail.Attachment($file) 
    $message.Subject = $subject 
    $message.Body = $body 
    $message.Attachments.Add($att) 
    $To = "to@domain.com" 
    $message.To.Add($to) 
    $message.From = $from 
             
    #Creating SMTP server object 
        
    
    #Sending email 
    $smtp.Send($message) 
    Write-Host "Email has been Sent!" 
  } 
  catch [System.Exception] 
  { 
    Write-Host "Mail Sending Error:" $_.Exception.Message -ForegroundColor Red 
  } 
} 
function Get-Cert($computer){ 
  $ro=[System.Security.Cryptography.X509Certificates.OpenFlags]"ReadOnly" 
  $lm=[System.Security.Cryptography.X509Certificates.StoreLocation]"LocalMachine" 
  $store=new-object System.Security.Cryptography.X509Certificates.X509Store("\\$computer\My",$lm) 
  $store.Open($ro) 
  $store.Certificates 
} 
 
function Get-RootCert($computer){ 
  $ro=[System.Security.Cryptography.X509Certificates.OpenFlags]"ReadOnly" 
  $lm=[System.Security.Cryptography.X509Certificates.StoreLocation]"LocalMachine" 
  $store=new-object System.Security.Cryptography.X509Certificates.X509Store("\\$computer\root",$lm) 
  $store.Open($ro) 
  $store.Certificates 
} 
 
 
$Servers = @("Server1","Server2") 
$datestring = (Get-Date).ToString("s").Replace(":","-")  
$file = "c:\temp\Certificates-$env:COMPUTERNAME-$datestring.csv"  
$Databases = @(); 
foreach($Server in $Servers) 
{ 
  $Certs = Get-Cert($Server) 
  foreach($Cert in $Certs) 
  { 
    $FriendlyName = $cert.FriendlyName 
    $Thumbprint = $Cert.Thumbprint 
    $Issuer = $Cert.Issuer 
    $Subject = $Cert.Subject 
    $SerialNumber = $Cert.SerialNumber 
    $NotAfter = $Cert.NotAfter 
    $NotBefore = $Cert.NotBefore 
    $DnsNameList = $cert.DnsNameList 
    $Version = $cert.Version 
  
    $DB = New-Object PSObject 
    Add-Member -input $DB noteproperty 'ComputerName' $Server 
    Add-Member -input $DB noteproperty 'FriendlyName' $FriendlyName 
    Add-Member -input $DB noteproperty 'DnsNameList' $DnsNameList 
    Add-Member -input $DB noteproperty 'ExpirationDate' $NotAfter 
    Add-Member -input $DB noteproperty 'IssueDate' $NotBefore 
    Add-Member -input $DB noteproperty 'Thumbprint' $Thumbprint 
    Add-Member -input $DB noteproperty 'Issuer' $Issuer 
    Add-Member -input $DB noteproperty 'Subject' $Subject 
    Add-Member -input $DB noteproperty 'SerialNumber' $SerialNumber 
    $Databases += $DB 
   } 
    
  $RootCerts = Get-RootCert($Server) 
  foreach($Cert in $RootCerts) 
  { 
    $FriendlyName = $cert.FriendlyName 
    $Thumbprint = $Cert.Thumbprint 
    $Issuer = $Cert.Issuer 
    $Subject = $Cert.Subject 
    $SerialNumber = $Cert.SerialNumber 
    $NotAfter = $Cert.NotAfter 
    $NotBefore = $Cert.NotBefore 
    $DnsNameList = $cert.DnsNameList 
    $Version = $cert.Version 
  
    $DB = New-Object PSObject 
    Add-Member -input $DB noteproperty 'ComputerName' $Server 
    Add-Member -input $DB noteproperty 'FriendlyName' $FriendlyName 
    Add-Member -input $DB noteproperty 'DnsNameList' $DnsNameList 
    Add-Member -input $DB noteproperty 'ExpirationDate' $NotAfter 
    Add-Member -input $DB noteproperty 'IssueDate' $NotBefore 
    Add-Member -input $DB noteproperty 'Thumbprint' $Thumbprint 
    Add-Member -input $DB noteproperty 'Issuer' $Issuer 
    Add-Member -input $DB noteproperty 'Subject' $Subject 
    Add-Member -input $DB noteproperty 'SerialNumber' $SerialNumber 
    $Databases += $DB 
   } 
    
} 
# $Databases | Out-GridView 
$Databases | Sort FriendlyName | Export-Csv -Path $file -NoTypeInformation -Append -Force 
SendMail "Email Subject " "Body : Server Certificates" $file  
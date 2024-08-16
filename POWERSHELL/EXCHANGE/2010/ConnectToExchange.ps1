## EWS Managed API Connect Module Script written by Glen Scales
## Requires the EWS Managed API and Powershell V2.0 or greator

## Load Managed API dll
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"

## Set Exchange Version
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1

## Create Exchange Service Object
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)

## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials

#Credentials Option 1 using UPN for the windows Account
$creds = New-Object System.Net.NetworkCredential("user@domain.com","password") 
$service.Credentials = $creds    

#Credentials Option 2
#service.UseDefaultCredentials = $true

## Choose to ignore any SSL Warning issues caused by Self Signed Certificates

[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use

#CAS URL Option 1 Autodiscover
$service.AutodiscoverUrl("email@domain.com",{$true})
"Using CAS Server : " + $Service.url 
 
#CAS URL Option 2 Hardcoded

#$uri=[system.URI] "https://casservername/ews/exchange.asmx"
#$service.Url = $uri  

## Optional section for Exchange Impersonation

#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, "email@domain.com")
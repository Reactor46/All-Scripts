##Build-SafeControlFromEntry
##Author:adamsor; adamsorenson.com
##Version: 1.0
 
###Example###
##$ULS = "The type Contoso.Customizations.WebPartControl, Contoso.WebParts, Version=1.0.0.0, Culture=neutral, PublicKeyToken=b03f5f8f11d50a3a could not be found or it is not registered as safe."
##$ULS | Build-SafeControlFromEntry
 
 
Function Build-SafeControlFromEntry
{
    param([Parameter(ValueFromPipeline=$true)]$string)
     
    #This code calls to a Microsoft web endpoint to track how often it is used. 
    #No data is sent on this call other than the application identifier
    Add-Type -AssemblyName System.Net.Http
    $client = New-Object -TypeName System.Net.Http.Httpclient
    $cont = New-Object -TypeName System.Net.Http.StringContent($null,[system.text.encoding]::UTF8,"application/json")
    $tsk = $client.PostAsync("https://msapptracker.azurewebsites.net/api/Hits/4b89c98c-473b-4fe3-8f8e-5552f13b0e30",$cont)
    #if you want to make sure the call completes, add this to the end of your code
    #$tsk.Wait()
     
    $wcSafeControl = $null
    $typeSafeControl = $null
     
    [int]$indexc = $string.IndexOf('The type')
    $indexTc = $indexc + 9
    $string2 = $string.Substring($indexTc)
    [int]$indexend = $string2.IndexOf("could not be found or it is not registered as safe.")
    $string3 = $string2.Substring(0, $indexend)
    $fcoma = $string3.IndexOf(",")
    $nt = $string3.Substring(0,$fcoma)
 
    $ASN = ($string3.Substring($fcoma+2)).Trim()
 
    $type = $nt.Substring($nt.LastIndexOf(".")+1)
    $namespace = $nt.Substring(0, $nt.LastIndexOf("."))
     
    If($ASN -ne $null -and $namespace -ne $null -and $type -ne $null)
    {
        $wcSafeControl = '<SafeControl Assembly="'+$ASN+'" Namespace="'+$namespace+'" TypeName="*" Safe="True" />'
        $typeSafeControl = '<SafeControl Assembly="'+$ASN+'" Namespace="'+$namespace+'" TypeName="'+$type+'" Safe="True" />'
        Write-host "SafeControls build successfully.  Use either Wildcard or Explicit in all of the web.configs from this web application" -ForegroundColor Green
 
        Write-host "Wildcard SafeControl" -ForegroundColor Yellow
        $wcSafeControl
 
        Write-Host "Explicit type SafeControl" -ForegroundColor Yellow
        $typeSafeControl
    }
    Else
    {
        Write-Host "Unable to find safe control from ULS entry.  Please look for 9tmwc for event ID.  Check KB" -ForegroundColor Red
        Write-Host "Example:" -ForegroundColor Red
        write-host "The type Contoso.Customizations.WebPartControl, Contoso.WebParts, Version=1.0.0.0, Culture=neutral, PublicKeyToken=b03f5f8f11d50a3a could not be found or it is not registered as safe."
    }
     
}
function Get-IISProperty {  
<#    
.SYNOPSIS    
    Retrieves IIS properties for Virtual and Web Directories residing on a server. 
.DESCRIPTION  
    Retrieves IIS properties for Virtual and Web Directories residing on a server. 
.PARAMETER name 
    Name of the IIS server you wish to query.  
.PARAMETER UseDefaultCredentials  
    Use the currently authenticated user's credentials    
.NOTES    
    Name: Get-iisProperties 
    Author: Marc Carter 
    DateCreated: 18Mar2011          
.EXAMPLE    
    Get-iisProperties -server "localhost"  
      
Description  
------------  
Returns IIS properties for Virtual and Web Directories residing on a server.  
#>  
<# 
[cmdletbinding(  
    DefaultParameterSetName = 'server',  
    ConfirmImpact = 'low'  
)]
#>  
Param(  
    [Parameter(  
        Mandatory = $True,  
        Position = 0,  
        ParameterSetName = '',  
        ValueFromPipeline = $True)]  
        [string][ValidatePattern(".{2,}")]$server 
)  
    Begin{       
        $error.clear() 
        $server = $server.toUpper() 
        $array = @() 
    } 
 
    Process{ 
        #define ManagementObjectSearcher, Path and Authentication 
        $objWMI = [WmiSearcher] "Select * From IIsWebServer" 
        $objWMI.Scope.Path = "\\$server\root\microsoftiisv2" 
        $objWMI.Scope.Options.Authentication = [System.Management.AuthenticationLevel]::PacketPrivacy 
        $server 
 
        trap { 'An Error occured: {0}' -f $_.Exception.Message; break } 
 
        #Get System.Management.ManagementObjectCollection 
        $obj = $objWMI.Get() 
     
        #Iterate through each object 
        $obj | ForEach-Object {  
            $Identifier = $_.Name 
            [string]$adsiPath = "IIS://$server/"+$_.name 
            $iis = [adsi]$("IIS://$server/"+$_.name) 
            #Enum Child Items but only IIsWebVirtualDir & IIsWebDirectory 
            $iis.Psbase.Children | Where-Object { $_.SchemaClassName -eq "IIsWebVirtualDir" -or $_.SchemaClassName -eq "IIsWebDirectory" } | ForEach-Object { 
                $currentPath = $adsiPath+"/"+$_.Name 
                #Enum Subordinate Child Items  
                $_.Psbase.Children | Where-Object { $_.SchemaClassName -eq "IIsWebVirtualDir" } | Select-Object Name, AppPoolId, SchemaClassName, Path | ForEach-Object { 
                    $subIIS = [adsi]$("$currentPath/"+$_.name) 
                    foreach($mapping in $subIIS.ScriptMaps){ 
                        if($mapping.StartsWith(".aspx")){ $NETversion = $mapping.substring(($mapping.toLower()).indexOf("framework\")+10,9) } 
                    } 
                    #Define System.Object | add member properties 
                    $tmpObj = New-Object Object 
                    $tmpObj | Add-Member -MemberType noteproperty -Name "Name" -Value $_.Name 
                    $tmpObj | Add-Member -MemberType noteproperty -Name "Identifier" -Value $Identifier 
                    $tmpObj | Add-Member -MemberType noteproperty -Name "ASP.NET" -Value $NETversion 
                    $tmpObj | Add-Member -MemberType noteproperty -Name "AppPoolId" -Value $($_.AppPoolId) 
                    $tmpObj | add-member -MemberType noteproperty -Name "SchemaClassName" -Value $_.SchemaClassName 
                    $tmpObj | Add-Member -MemberType noteproperty -Name "Path" -Value $($_.Path) 
                     
                    #Populate Array with Object properties 
                    $array += $tmpObj 
                } 
            } 
        } 
    }#End process 
    End{ 
        #Display results 
        $array | ft -AutoSize 
    } 
}#End function
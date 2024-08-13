Import-Module -name GroupPolicy

$now = Get-date
$date = get-date -uformat "%Y_%m_%d_%I%M%p"

#------- Modify this line for csv file
$CSVpath = "C:\LazyWinAdmin\Logs\GPO_Settings_$date.csv"

#------- Modify this line with name of computer policy you are looking for
$Setting = "Configure user Group Policy loopback processing mode"

# ------ Next four lines for emailing of report
$SMTPServer = "mailgateway.Contoso.corp"
$From = "GPOReports@creditone.com"
$To = "john.battista@creditone.com"
$Subject = "Has $Setting $Date"

[array]$Report = @()

$GPOs = Get-GPO -all | Sort-Object Displayname

foreach ($GPO in $GPOs) 
{
    $XMLReport = Get-GPOReport -GUID $($GPO.id) -ReportType xml
    $XML = [xml]$XMLReport
    
	$Types = @("Computer")
	
    Foreach ($Type in $Types)
	{
		Write-Host "Processing GPO $($GPO.DisplayName)"
        $ExtArray = $xml.gpo.$Type.ExtensionData | foreach-Object -process {$_.name}
		
        if ($ExtArray -ne $Null){$Ext = [array]::IndexOf($ExtArray, 'Registry')}
        
        $ExtCount = $ExtArray.count
        #write-host "Extension count is $ExtCount"
        #write-host "EXT is $Ext"
	    
        if (($Ext -eq -1) -or ($Ext -eq $Null) -or ($ExtCount -eq 0))
	    {
	        #Write-Host "GPO Name $($GPO.Displayname) is Null"
	        Clear-Variable ExtArray
	        Clear-Variable EXT
	    }
        
        Else
        {	
            #write-host "Has Reg Settings"
            if ($ExtCount -eq 1){$SwitchVal = 1}
            Else {$SwitchVal = 2}
        
            Switch ($SwitchVal)
            {
                1{$Extension = $xml.gpo.$Type.ExtensionData.Extension.policy}
                2{$Extension = $xml.gpo.$Type.ExtensionData[$Ext].Extension.policy}
            }

            $ValArray = $Extension | ForEach-Object -process {$_.name}
            if ($ValArray -ne $Null){$Val = [array]::IndexOf($ValArray, $Setting)}
            $ValCount = $ValArray.count
            #write-host "Val count is $Valcount"
            #write-host "Val is $Val"

            if (($Val -eq -1) -or ($val -eq $Null) -or ($ValCount -eq 0))
	        {
	            #Write-Host "GPO Name $($GPO.Displayname) Doesn't have setting"
	            Clear-Variable Valarray
	            Clear-Variable Val
            }

            Else
            {

               if (($ValCount -eq 1) -and ($Switchval -eq 1)){$SwitchVal2 = 1}
               elseif (($ValCount -eq 1) -and ($Switchval -eq 2)){$SwitchVal2 = 3}
               elseif (($Valcount -gt 1) -and ($Switchval -eq 1)){$SwitchVal2 = 2}
               else {$SwitchVal2 = 4}

               Switch ($SwitchVal2)
               {
                    1{
                        $PolicyState = $xml.gpo.$Type.ExtensionData.Extension.policy.state
                        $Extension = $xml.gpo.$Type.ExtensionData.Extension.policy.Name
                        }
                    2{
                        $PolicyState = $xml.gpo.$Type.ExtensionData.Extension.policy[$Val].state
                        $Extension = $xml.gpo.$Type.ExtensionData.Extension.policy[$Val].Name
                        }
                    3{
                        $PolicyState = $xml.gpo.$Type.ExtensionData[$Ext].Extension.policy.state
                        $Extension = $xml.gpo.$Type.ExtensionData[$Ext].Extension.policy.name
                        }
                    4{
                        $PolicyState = $xml.gpo.$Type.ExtensionData[$Ext].Extension.policy[$Val].state
                        $Extension = $xml.gpo.$Type.ExtensionData[$Ext].Extension.policy[$Val].name
                        }
               }

	           #Write-Host "GPO Name $($GPO.Displayname) is Not Null"
	           #write-host "Extension is $Extension"	            

               if ($Extension -eq $Setting)
	           {
	                $Report += New-Object PSObject -Property @{
	                    'GPO Name' = $xml.gpo.name
	                    'GPO Type' = $Type
				        'Has Setting' = $True
	                    'Policy Set State' = $PolicyState
	                    'Setting Enabled' = $xml.gpo.$Type.enabled
	                    }
	                Clear-variable Ext
	                Clear-Variable ExtArray
                    Clear-Variable Valarray
	                Clear-Variable Val
	            }    
	                Clear-Variable Extension
                    clear-variable Policystate
            }
        }
	}
}


$HTMLHeader = @"
 <style>
 TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
 TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
 TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
 </style>
"@

$SMTPMessage = @{
To = $To
From = $From
Subject = $Subject
Smtpserver = $SMTPServer
Attachments = $CSVpath
}

$CSVReport = $Report | Export-CSv -path $CSVpath

$HTMLReport = $Report | ConvertTo-HTML -Head $HTMLHeader

Send-MailMessage @SMTPMessage -Body ($HTMLReport | Out-String) -bodyashtml

Remove-Variable Report

Remove-Item -path $CSVpath
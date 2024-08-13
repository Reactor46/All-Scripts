<#    
.SYNOPSIS    
    
  Get all computer objects from Active Directory with "server" in the operating system name and seperate them by Windows version 
        
.COMPATABILITY     
     
  Tested on PS v4. 
      
.EXAMPLE  
  PS C:\> GetServerOS.ps1  
  All options are set as variables in the GLOBALS section so you simply run the script.  
  
.NOTES    
        
  NAME:       GetServerOS.ps1    
    
  AUTHOR:     Brian D. Arnold    
    
  CREATED:    3/20/14   
    
  LASTEDIT:   6/27/14   
#>

# Import SQLPS module which contains invoke-sqlcmd.
#Import-Module ActiveDirectory

###################
##### GLOBALS #####
###################

# Text to be displayed on each line before the results
$txt_groupA = "Server 2012 Standard........"
$txt_groupB = "Server 2012 Datacenter......"
$txt_groupC = "Server 2012 R2 Standard....."
$txt_groupD = "Server 2012 R2 Datacenter..."
$txt_groupE = "Server 2016 Standard........"
$txt_groupF = "Server 2016 Datacenter......"
$txt_groupG = "Total Windows Servers......."

################
##### MAIN #####
################

# Get all AD computers with "server" in the opertaingsystem property
$Servers = Get-Content -Path .\Configs\Contoso.txt
ForEach($srv in $Servers){
$results = Get-ComputerInfo -ComputerName $srv | Select ComputerName, OSName
}

$groupA = $results.OSName -like "*Server 2012 Standard*" 
Write-Host -NoNewline $txt_groupA
$groupA.count

$groupB = $results.OSName -like "*Server 2012 Datacenter*" 
Write-Host -NoNewline $txt_groupB
$groupB.count

$groupC = $results.OSName -like "*Server 2012 R2 Standard*" 
Write-Host -NoNewline $txt_groupC
$groupC.count

$groupD = $results.OSName -like "*Server 2012 R2 Datacenter*" 
Write-Host -NoNewline $txt_groupD
$groupD.count

$groupE = $results.OSName -like "*Server 2016 Standard*" 
Write-Host -NoNewline $txt_groupE
$groupE.count

$groupF = $results.OSName -like "*Server 2016 Datacenter*" 
Write-Host -NoNewline $txt_groupF
$groupF.count

$groupG.count | Write-Host -ForegroundColor Green
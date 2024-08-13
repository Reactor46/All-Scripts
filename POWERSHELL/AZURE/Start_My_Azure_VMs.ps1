<#

.DESCRIPTION


.NOTES
	Author: Peter Selch Dahl - Installers A/S
	Last Updated: 4/14/2014  
#>

workflow Start_My_Azure_VMs
{   
      param()

       $MyConnection = "My MSDN Azure Account Connection"
       $MyCert = "MSDN Azure Certifcate"


    # Get the Azure Automation Connection
    $Con = Get-AutomationConnection -Name $MyConnection
    if ($Con -eq $null)
    {
        Write-Output "Connection entered: $MyConnection does not exist in the automation service. Please create one `n"   
    }
    else
    {
        $SubscriptionID = $Con.SubscriptionID
        $ManagementCertificate = $Con.AutomationCertificateName
       
    }   

    # Get Certificate & print out its properties
    $Cert = Get-AutomationCertificate -Name $MyCert
    if ($Cert -eq $null)
    {
        Write-Output "Certificate entered: $MyCert does not exist in the automation service. Please create one `n"   
    }
    else
    {
        $Thumbprint = $Cert.Thumbprint
    }

        #Set and Select the Azure Subscription
         Set-AzureSubscription `
            -SubscriptionName "My Azure Subscription" `            -Certificate $Cert `            -SubscriptionId $SubscriptionID `

        #Select Azure Subscription
         Select-AzureSubscription `
            -SubscriptionName "My Azure Subscription"

    Write-Output "-------------------------------------------------------------------------"

       Write-Output "Starting the Domain Controllers.."

       # Please type the name of your Domain Controllers

      Start-AzureVM -ServiceName "SCOMDC" -Name "SCOMDC"
      #Start-AzureVM -ServiceName "SCOMDC" -Name "SCOMDC"
      #Start-AzureVM -ServiceName "SCOMDC" -Name "SCOMDC"
      #Start-AzureVM -ServiceName "SCOMDC" -Name "SCOMDC"
      
      Write-Output "Sleeping for 60 seconds. Waiting for Domain Controllers to come online..."
      Start-Sleep -Seconds 60

       Write-Output "-------------------------------------------------------------------------"
       Write-Output "Starting the Virtual Machines NOW!"

       Get-AzureVM | select name | ForEach-Object {
        $StopOutPut = Start-AzureVM -ServiceName $_.Name -Name $_.Name
           Write-Output "Starting.... :  $_.Name "
        #   Write-Output $StopOutPut
           
           }

       Write-Output "-------------------------------------------------------------------------"
 
}
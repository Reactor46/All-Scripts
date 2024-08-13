<#
.Synopsis
   This script will produce a list of all open endpoints configured on VMs running in an Azure subscription
.DESCRIPTION
   This will look for subscriptions - if none are found you will be prompted to add an account.
      This will print to your screen for ever VM in the subscriptin their confirured Azure Endpoints
.INPUTS
   None
.OUTPUTS
   Output from this cmdlet - a printed list on your screen of open endpoints for all VMs in a subscription
.NOTES
   This script idea came from wanting to prevent portal users from opening endpoints in a subscription that has a company VPN and a configured VNET using private IP adderss
   We want to aduit any endpoints that portal owners may open
.EXAMPLE #1
   Find-AllEndPointsAll

#>
function Find-AzureEndPointsAll
{
    [CmdletBinding()]
    Param
    (
 
    )
    Begin
    {
        #Find subscriptions and ask to validate or add Azure User

            #Gets the subscription listed for the user
            $allsubs = Get-AzureSubscription
            Clear-Host

            if ($allsubs)
            {
                
            }
            Else
            {
                Write-Host "No Subscriptions were found please add an Azure account that is associated with a subscritpion you wish to audit."
                Read-Host "Press Enter to Continue"

                #Prompt user to log in an add account wtih subscriptions
                Add-AzureAccount

            }
                #Get a list of Subscriptions 
                $allsubslist = Get-AzureSubscription

                Write-Host "The following Azure Subscriptions will be scanned for EndPoints"
                Write-Host ""

                foreach($allsubslist in $allsubslist)
                    {
                        Write-Host $allsubslist.SubscriptionName
                    }
                        Write-Host ""
    }
    Process
    {
        #Sets Subscription count to 0
        $count = 0
        
        # This will capture the users current default subscription name
        $currentdefaulsub = Get-AzureSubscription -Default
        $currentsubname = $currentdefaulsub.SubscriptionName

        #Get the list of Subscriptions to scan
        $allsubs = Get-AzureSubscription

        #Will run the following for every subscrtipion found       
        foreach ($allsub in $allsubs)
            {
                #sets a subscription to use for this pass
                ##Select-AzureSubscription -Default $allsubs.
                Select-AzureSubscription $allsubs.SubscriptionName[$count]

                    #Gets all the Services in the subscription
                    $allservices =  Get-AzureVM

                    Write-Host "VMs Found in" ($allsubs.SubscriptionName[$count])
                   
                    foreach ($allservices in $allservices)
                        {
                            #looks up the VM name and provides data on status and endpoints
                            Get-AzureVM -ServiceName $allservices.servicename | ft -Property Name, DNSName, IpAddress, Powerstate -AutoSize
                            Get-AzureVM -ServiceName $allservices.servicename | Get-AzureEndpoint | ft -Property localport, name, port, protocol, Vip -AutoSize
                        }
             
                # Increments the subscription count
                $count= $count + 1
            }
    }
    End
    {
        #Sets the users default subscription back to what it as befor the script 
        Select-AzureSubscription -Default $currentsubname
         
    }
}


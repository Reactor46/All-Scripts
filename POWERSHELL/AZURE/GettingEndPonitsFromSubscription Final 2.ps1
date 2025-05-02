<#
.Synopsis
   This script will produce a list of all open endpoints configured on VMs running in an Azure subscription
.DESCRIPTION
   You will be promted for a susscription name or you can leave blank to user your already set default subscription
   This will print to your screen for ever VM in the subscriptin their confirured Azure Endpoints
.INPUTS
   Inputs to this cmdlet - you need to give it a subscritpion name or it will use your default subscription
.OUTPUTS
   Output from this cmdlet - a printed list on your screen of open endpoints for all VMs in a subscription
.NOTES
   This script idea came from wanting to prevent portal users from opening endpoints in a subscription that has a company VPN and a configured VNET using private IP adderss
   We want to aduit any endpoints that portal owners may open
.EXAMPLE #1
   Find-AllEndPoints
.EXAMPLE #2
   Find-AllEndPoints -SubscriptionaName <your subscription name>
#>
function Find-AzureEndPoints
{
    [CmdletBinding()]
    Param
    (
        #Param to get the subscription to use
                [Parameter(Mandatory=$true,
                    HelpMessage = "Enter Subscription Name or leave empty to use default subscription (names are case senitive)")]
                [AllowEmptyString()]
                [string]$SubscriptionaName
    )
    Begin
    {
        #Will change default subscription to what user entered
        if ($subscriptionaname)
          {
              # This will capture the user current default subscription name
              $currentdefaulsub = Get-AzureSubscription -Default
              $currentsubname = $currentdefaulsub.SubscriptionName

              # This will set the default subscription to what the user inputed - if not left blank
              Set-AzureSubscription -DefaultSubscription $subscriptionaname
          }
    }
    Process
    {
         #Gets all the Services in the subscription
         $allservices =  Get-AzureService

                foreach ($allservices in $allservices)
            {
                #looks up the VM name and provides data on status and endpoints
                Get-AzureVM -ServiceName $allservices.servicename | ft -Property Name, DNSName, IpAddress, Powerstate -AutoSize
                Get-AzureVM -ServiceName $allservices.servicename | Get-AzureEndpoint | ft -Property localport, name, port, protocol, Vip -AutoSize
            }
    }
    End
    {
        #Sets the users default subscription back to what it as befor the script - assuming it was changed/modified
        if ($SubscriptionaName)
          {
            Set-AzureSubscription -DefaultSubscription $currentsubname
          }
    }
}


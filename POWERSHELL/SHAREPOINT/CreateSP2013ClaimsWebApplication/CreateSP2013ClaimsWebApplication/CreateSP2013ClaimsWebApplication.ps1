$ver = $host | select version
if($Ver.version.major -gt 1) {$Host.Runspace.ThreadOptions = "ReuseThread"}
if(!(Get-PSSnapin Microsoft.SharePoint.PowerShell -ea 0))
{
Write-Progress -Activity "Loading Modules" -Status "Loading Microsoft.SharePoint.PowerShell"
Add-PSSnapin Microsoft.SharePoint.PowerShell
}

Write-Progress -Activity "Creating Web Application" -Status "Setting Variables"
##
#Set Individual Web App Variables
##

#This is the Web Application URL
$WebApplicationURL = "http://Contoso.com"

#This is the Display Name for the SharePoint Web Application
$WebApplicationName = "Contoso SharePoint Site"

#This is the Content Database for the Web Application
$ContentDatabase = "Contoso_ContentDB"

##
#Set Common Variables
##

#This is the Display Name for the Application Pool
$ApplicationPoolDisplayName = "SharePoint App Pool"

#This is identity of the Application Pool which will be used (Domain\User)
$ApplicationPoolIdentity = "Contoso\SPAppPool"

#This is the password of the Appliation Pool account which will be used
$ApplicationPoolPassword = "Pass@word1"

#This is the Account which will be used for the Portal Super Reader Account
$PortalSuperReader = "i:0#.w|Contoso\SuperReader"

#This is the Account which will be used for the Portal Super User Account
$PortalSuperUser = "i:0#.w|Contoso\SuperUser"


Write-Progress -Activity "Creating Web Application" -Status "Loading Functions"

##
#Create Functions
##

Function CreateClaimsWebApp($WebApplicationName, $WebApplicationURL, $ContentDatabase, $HTTPPort)
{
    #AppPoolUsed is set when calling the ValidateAppPool function. This will be true if the application pool is already running SharePoint web appplications
    #If the application pool is already being used in web applications, the syntax for New-SPWebApplication changes
    if($AppPoolUsed -eq $True)
    {
        #Create the web application, assign it to the WebApp variable.  The WebApp variable will be used to set object cache user accounts
        Write-Progress -Activity "Creating Web Application" -Status "Using Application Pool With Existing Web Applications"
        Set-Variable -Name WebApp -Value (New-SPWebApplication -ApplicationPool $ApplicationPoolDisplayName -Name $WebApplicationName -url $WebApplicationURL -port $HTTPPort -DatabaseName $ContentDatabase -HostHeader $hostHeader -AuthenticationProvider (New-SPAuthenticationProvider)) -Scope Script
        
        #Call the SetObjectCache function, which sets the object cache.
        Write-Progress -Activity "Creating Web Application" -Status "Configuring Object Cache Accounts"
        SetObjectCache
        
    }
    else
    {
        #Create the web application, assign it to the WebApp variable.  The WebApp variable will be used to set object cache user accounts
        Write-Progress -Activity "Creating Web Application" -Status "Using Application Pool With No Existing Web Applications"
        Set-Variable -Name WebApp -Value (New-SPWebApplication -ApplicationPool $ApplicationPoolDisplayName -ApplicationPoolAccount $AppPoolManagedAccount.Username -Name $WebApplicationName -url $WebApplicationURL -port $HTTPPort -DatabaseName $ContentDatabase -HostHeader $hostHeader -AuthenticationProvider (New-SPAuthenticationProvider)) -Scope Script
        
        #Call the SetObjectCache function, which sets the object cache.
        Write-Progress -Activity "Creating Web Application" -Status "Configuring Object Cache Accounts"
        SetObjectCache
        
    }
}

Function ValidateURL($WebApplicationURL)
{
    #Find out if a web application with the target URL exists
    if(get-spwebapplication $WebApplicationURL -ErrorAction SilentlyContinue)
    {
        #If a web application with the specifid URL already exists, wait 5 seconds and exit
        Write-Progress -Activity "Creating Web Application" -Status "Aborting Process Due To URL Conflict"
        Write-Host "Aborting: Web Application $WebApplicationURL Already Exists" -ForegroundColor Red
        sleep 5

        #Setting the CriticalError value to $True results in the script to not create anything
        Set-Variable -Name CriticalError -Value $True
    }
    
    #If the WebApplicationURL passed is not already a SharePoint web application, find out if it starts with HTTP or HTTPS
    elseif($WebApplicationURL.StartsWith("http://"))
        {
            #If the string starts with http://, and not https://, trim the protocol from the URL.  Set the host as the host header
            Set-Variable HostHeader -Value ($WebApplicationURL.Substring(7)) -Scope Script

            #If we're using HTTP, use port 80
            Set-Variable -Name HTTPPort -Value "80" -Scope Script
        }
        elseif($WebApplicationURL.StartsWith("https://"))
        {
            #If the string starts with https://, and not http://, trim the protocol from the URL.  Set the host as the host header
            Set-Variable HostHeader -Value ($WebApplicationURL.Substring(8)) -Scope Script

            #If we're using HTTPS, use port 443
            Set-Variable -Name HTTPPort -Value "443" -Scope Script
        }
}

Function ValidateAppPool($AppPoolName, $WebApplicationURL)
{
    #Change the ErrorActionPreference to SilentlyContinue while preserving the original value in a temporary variable
    #Failing to do this will result in error messages being displayed if Get-WebAppPoolState does not return an object.  The script would still continue
    $CurrentErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "SilentlyContinue"

    #Check to see if an application pool with the name passed by the AppPoolName variable already exists, assign this to a variable.
    #This variable will be used in order to determine if the application pool exists, but is not part of SharePoint
    $TestAppPool = Get-WebAppPoolState $AppPoolName

    #If we have a SharePoint application pool with the value passed by the AppPoolName variable, find out if there are any sites using that app pool
    #This changes the syntax used with New-SPWebApplication
    if(Get-SPServiceApplicationPool $AppPoolName)
    {
        #Return all application pools used by all web applications
        $AppPools = Get-SPWebApplication | select ApplicationPool

        #Providing there is more than one application pool, find out what their names are
        if($AppPools)
        {
            foreach($Pool in $AppPools)
            {
                #Get The application pool display name for each application pool returned
                [Array]$Poolchild = $Poolchild += ($Pool.ApplicationPool.DisplayName)

                #If any application pool matches the value passed by ApplicationPoolDisplayName, set AppPoolUsed to True
                #This is referenced in the CreateClaimsWebApp function
                if($Poolchild.Contains($ApplicationPoolDisplayName))
                {
                    Set-Variable -Name AppPoolUsed -Value $True -Scope Script
                }

                #If the application pool display name does not match the value passed by ApplicationPoolDisplayName, set AppPoolUsed to False
                #This is referenced in the CreateClaimsWebApp function
                else
                {
                    Set-Variable -Name AppPoolUsed -Value $False -Scope Script
                }
            }
        }
        
        #Since this is a SharePoint Application Pool, set the AppPool value to the the SPServiceApplicationPool object returned
        Set-Variable -Name AppPool -Value (Get-SPServiceApplicationPool $AppPoolName) -scope Script

        #Set the AppPoolManagedAccount variable to the name of the managed acount used by the application pool returned
        #AppPoolManagedAccount is used in the CreateClaimsWebApp function if the application pool does not have existing web applications that are using it
        Set-Variable -Name AppPoolManagedAccount -Value (Get-SPManagedAccount | ? {$_.username -eq ($AppPool.ProcessAccountName)}) -scope Script
    }

    #Check to see if the application pool is in IIS, but is not a SharePoint app pool
    elseif($TestAppPool)
    {
        #If the application pool exists in IIS and is not a SharePoint application pool, abort the script by setting CriticalError to True
        Write-Host "Aborting: Application Pool $AppPoolName already exists on the server and is not a SharePoint Application Pool `n`rWeb Application `"$WebApplicationURL`" will not be created" -ForegroundColor Red
        Set-Variable -Name CriticalError -Value $True
    }
    #If it's not a SharePoint app pool, and it doesn't exist in IIS, we have to create one
    elseif(!($TestAppPool))
    {
        #Find out if a managed account exists by calling the ValidateManagedAccount function
        validateManagedAccount $ApplicationPoolIdentity

        #If the managed account exists, create an application pool using the existing managed account
        if($ManagedAccountExists -eq $True)
        {
            #Set the AppPoolManagedAccount to the identity of the managed acocunt referenced by the ApplicationPoolIdentity variable
            Write-Host "Creating New App Pool using Existing Managed Account"
            Set-Variable -Name AppPoolManagedAccount -Value (Get-SPManagedAccount $ApplicationPoolIdentity | select username) -scope "Script"

            #Create a new SPServiceApplicationPool, assign that to the AppPool variable
            Set-Variable -Name AppPool -Value (New-SPServiceApplicationPool -Name $ApplicationPoolDisplayName -Account $ApplicationPoolIdentity) -scope "Script"
        }

        #If there is no managed account matching the account referenced by the ApplicationPoolIdentity, create it
        else
        {
            #Use the ApplicationPoolIdentity and ApplicationPoolPassword to create a credential object
            #This is necessary when creating a new managed account
            Write-Host "Creating New Managed Account And App Pool"
            $AppPoolCredentials = New-Object System.Management.Automation.PSCredential $ApplicationPoolIdentity, (ConvertTo-SecureString $ApplicationPoolPassword -AsPlainText -Force)

            #Create a new managed account, assign that to the AppPoolManagedAccount variable
            Set-Variable -Name AppPoolManagedAccount -Value (New-SPManagedAccount -Credential $AppPoolCredentials) -scope "Script"

            #Create a new application pool using the new managed account, assign this to the AppPool variable
            Set-Variable -Name AppPool -Value (New-SPServiceApplicationPool -Name $ApplicationPoolDisplayName -Account (get-spmanagedaccount $ApplicationPoolIdentity)) -scope "Script"
        }

    }
    
    #Return the ErrorActionPreference to the default value
    $ErrorActionPreference = $CurrentErrorActionPreference

}

Function ValidateManagedAccount($ApplicationPoolIdentity)
{
    #Find out if the manage account referenced by the AppPoolIdentity already exists
    #If it does, set ManagedAccountExists to True
    if(Get-SPManagedAccount $ApplicationPoolIdentity -ErrorAction SilentlyContinue)
    {
        Set-Variable -Name ManagedAccountExists -Value $True -Scope Script
    }
    #If it does not, set ManagedAccountExists to False
    else
    {
        Set-Variable -Name ManagedAccountExists -Value $False -Scope Script
    }
}

Function ClearScriptVariables
{
    #Set the ErrorActionPreference to SilentlyContinue
    #If this is not set, and the script variables referenced have not been set, an error message will be returned.  The script would still continue
    $CurrentErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "SilentlyContinue"

    #Remove the CriticalError variable
    Remove-Variable $CriticalError -ErrorAction SilentlyContinue
    $ErrorActionPreference = $CurrentErrorActionPreference
}
Function SetObjectCache
{
    #Set object cache user account properties based on the value of the parameters supplied
    $WebApp.Properties["portalsuperuseraccount"] = $PortalSuperUser
    $WebApp.Properties["portalsuperreaderaccount"] = $PortalSuperReader
     
    #Create a New Policy for the Super User
    $SuperUserPolicy = $WebApp.Policies.Add($PortalSuperUser, "Portal Super User Account")

    #Assign Full Control To the Super User
    $SuperUserPolicy.PolicyRoleBindings.Add($WebApp.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::FullControl))
    
    #Create a New Policy for the Super Reader
    $SuperReaderPolicy = $WebApp.Policies.Add($PortalSuperReader, "Portal Super Reader Account")
    
    #Assign Full Read to the Super Reader
    $SuperReaderPolicy.PolicyRoleBindings.Add($WebApp.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::FullRead))

    #Commit these changes to the web application
    $WebApp.Update()
}


##
#Script
##

#Call the ClearScriptVariables function to empty out varialbes that should be blank when the script executes.

ClearScriptVariables

#Validate the URL passed by calling the ValidateURL function
Write-Progress -Activity "Creating Web Application" -Status "Validating Web Application URL Variables"
ValidateURL $WebApplicationURL

#Validate the application pool variables by calling the ValidateAppPool function
Write-Progress -Activity "Creating Web Application" -Status "Validating Application Pool Variables"
ValidateAppPool $ApplicationPoolDisplayName $WebApplicationURL


#As long as CriticalError has not been set, create the web application using the variables passed.
if(!($CriticalError))
{
Write-Progress -Activity "Creating Web Application" -Status "Creating Claims-Based Web Application"
CreateClaimsWebApp $WebApplicationName $WebApplicationURL $ContentDatabase $HTTPPort
}

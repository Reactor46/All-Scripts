Function Connect-SPOnline {
  <#
  .SYNOPSIS
  This PowerShell function, Connect-SPOnline, connects an Office 365  SharePoint Online global administrator to their Administration Center site with PowerShell 3.0. You can elect to provide a credential password each time, or save it for automated connection reuse.  

  .DESCRIPTION
  This PowerShell function, Connect-SPOnline, connects an Office 365  SharePoint Online global administrator to their Administration Center site with PowerShell 3.0. You can elect to provide a credential password each time, or save it for automated connection reuse.
 
  .EXAMPLE
  Connect-SPOnline -User "craig@mysponlinesite.com" -Url "https://mysponlinesite-admin.sharepoint.com"

  This example will connect a user to SharePoint Online and will require a user to provide a password for the credential sent to SharePoint Online.

  .EXAMPLE
  Connect-SPOnline -User "craig@mysponlinesite.com" -Url "https://mysponlinesite-admin.sharepoint.com" -UseStoredCredentials

  This example will connect a user to SharePoint Online and will use a stored password for the credential sent to SharePoint Online. If this is the first time you are using the -Use StoredCredentials switch, you will be prompted to provide a password so it can be saved in the location defined by the StoredCredentialPath parameter.

  .EXAMPLE
  $SPOnlineParameters = @{ 
        User = "craig@mysponlinesite.com" 
        Url = "https://mysponlinesite-admin.sharepoint.com" 
        UseStoredCredentials = $true
  } 
  Connect-SPOnline @SPOnlineParameters

  This example uses the PowerShell 3.0 spalatting technique to pass parameters to the function and will connect a user to SharePoint Online using a stored password for the credential sent to SharePoint Online. If this is the first time you are using the -Use StoredCredentials switch, you will be prompted to provide a password so it can be saved in the location defined by the StoredCredentialPath parameter.

  .PARAMETER User
  Required. String. The username of a SharePoint Online global administrator who can access the SharePoint Online Administration Center site. 

  .PARAMETER Url
  Required. String. Specifies the URL of the SharePoint Online Administration Center site. 

  .PARAMETER UseStoredCredentials
  Optional. Switch. If specified, the function will attempt to use the secured locally stored password for the specified User to connect to SharePoint Online. If not specified, the user will be prompted to provide a credential password. The first time you run the function with this switch, you will be prompted for your credential password in order to save your credentials for reuse.

  .PARAMETER StoredCredentialPath
  Optional. String. Specifies the folder path where stored credential files are saved. The default path for this parameter is a 'StoredSPOnlineCredentials' directory within a user's $profile directory. Unless you wish to change the location of where credentials are stored, you don't need to use this parameter.

  .Notes
  Name: Connect-SPOnline
  Author: Craig Lussier
  Last Edit: March 2nd, 2013

  .Link
  http://www.craiglussier.com
  http://twitter.com/craiglussier
  http://social.technet.microsoft.com/profile/craig%20lussier/

  # Requires PowerShell Version 3.0
  #>
[CmdletBinding()]
    param (
        [Parameter(Position=0, Mandatory=$true)]
        [string]$User, #= "craiglussier@phoenixlab.onmicrosoft.com",

        [Parameter(Position=1, Mandatory=$true)]
        [string]$Url, # = "https://phoenixlab-admin.sharepoint.com",

        [Parameter(Position=2, Mandatory=$false)]
        [switch]$UseStoredCredentials,

        [Parameter(Position=3, Mandatory=$false)]
        [string]$StoredCredentialPath = (Split-Path $profile) + "\StoredSPOnlineCredentials"


    )

    Begin {

        Write-Verbose "Entering Begin Block: Connect-SPOnline"
        
        Write-Verbose "Determine PowerShell Host Major Version"
        $hostMajorVersion = $Host.Version.Major
        if($hostMajorVersion -lt 3) {
            
            Write-Error "The Connect-SPOnline function requires PowerShell Version 3. You are running PowerShell Version $hostMajorVersion. Exiting function."
            Exit

        }
        else {

            Write-Verbose "Current PowerShell Version is $hostMajorVersion"

        }
                             
        Write-Verbose "Determine the 'bitness' of the current PowerShell Process"
        $bitness = $null        
        if([Environment]::Is64BitProcess) {            
            $bitness = "64"
        }
        else {
            $bitness = "32"
        }
        Write-Verbose "The current PowerShell process is $bitness-bit"

        $ModuleName = "Microsoft.Online.SharePoint.PowerShell"
        Write-Verbose "Determine if the $ModuleName module is loaded"
        try {

            if(-not(Get-Module -name $name)) 
            { 

                Write-Verbose "Load the SharePoint Online Module"
                Import-Module -Name $name -DisableNameChecking -ErrorAction Stop
                Write-Verbose "Successfully loaded the SharePoint Online Module"

            }
            else {

                Write-Verbose "The SharePoint Online Module was already loaded in this PowerShell session prior to executing this function."
            
            }

        }
        catch {

            Write-Error "There was an error while attempting to load the SharePoint Online Module. Exiting function. Please note that the current PowerShell process is $bitness-bit. There are two versions of the SharePoint Online Management Shell which includes the $ModuleName module - a 32-bit version and 64-bit version. Please ensure that you have the correct version installed to run this function within this $bitness-bit PowerShell process. The SharePoint Online Management Shell is available for download at http://www.microsoft.com/en-us/download/details.aspx?id=35588."
            Write-Error $_
            Exit

        }

        Write-Verbose "Leaving Begin Block: Connect-SPOnline"
    }

    Process {

        Write-Verbose "Entering Process Block: Connect-SPOnline"
    
        if($UseStoredCredentials) {

            Write-Verbose "The stored password for will be used to connect to SharePoint Online for the $User user credential." 
            
            Write-Verbose "Remove Trailing Directory Slash \ if it exists in the StoredCredentialPath function parameter."
            $StoredCredentialPath = $StoredCredentialPath.TrimEnd("\")

            Write-Verbose "Constructing path to stored credentials for user $User"
            $fileName = "$StoredCredentialPath\$User-Connect-SPOnline-Credentials.txt"
            

            Write-Verbose "Check if stored credentials exist"
            if(!(Test-Path $fileName)) {

                Write-Warning "Stored credentials do not exist for $User - Stored credentials will be created." 
                Write-Warning "You will be prompted for a password. This step will only occur the first time you use this functionality."

                Write-Verbose "Check to see if the stored credential path exists"
                if(!(Test-Path $StoredCredentialPath)) {

                    try {

                        Write-Verbose "Creating directory to store credentials"
                        New-Item -Path $StoredCredentialPath -ItemType Directory | Out-Null

                    }
                    catch {

                        Write-Error "An error occurred while attemptiong to create the directory $StoredCredentialPath. Exiting function."
                        Write-Error $_
                        Exit

                    }

                }
                else {

                    Write-Verbose "Stored credential path exists."

                }

                try {

                    Write-Verbose "Prompt for password"
                    $credential = Get-Credential $User               

                    Write-Verbose "Save credentials in $fileName"
                    $credential.Password | ConvertFrom-SecureString | Set-Content $fileName -Force
                     
                }
                catch {

                    Write-Error "An error occurred while capturing credentials and saving the credentials in $fileName. Exiting function."
                    Write-Error $_
                    Exit

                }

            }
            else {

                Write-Verbose "Stored Credentials exist"

            }

            $password = $null

            try {

                Write-Verbose "Retrieving stored password from file."
                $password = Get-Content $fileName | ConvertTo-SecureString 
                Write-Verbose "Successfully retrieved stored password from file."

            }
            catch {

                Write-Error "An error occurred while retrieving your stored secure credentials. Exiting function."
                Write-Error $_
                Exit

            }

            try {

                Write-Verbose "Creating credential object for $User with stored password to send to SharePoint Online"
                $credential = New-Object System.Management.Automation.PSCredential $User, $password

                Write-Verbose "Connecting to SharePoint Online"
                Connect-SPOService -Url $Url -Credential $credential
                Write-Verbose "Connected to SharePoint Online"

            }
            catch {

                Write-Error "An error occurred while attempting to connect to SharePoint Online. Exiting function. Your issue is either with your Internet connection, the SharePoint Online Administration site URL you specified is incorrect, the credentials for the SharePoint Online Administration site you specified are incorrect or the SharePoint Online system itself is experiencing a problem. Please read the error message below carefully for clues on how to resolve your issue."
                Write-Error $_
                Exit

            }

        }
        else {

            Write-Verbose "User elected to be prompted for credential password." 

            try {

                Write-Warning "Provide a credential password for $User for site $URL. It is requried to connect to SharePoint Online."
                $credential = Get-Credential $User 

                Write-Verbose "Connecting to SharePoint Online"
                Connect-SPOService -Url $Url -Credential $credential
                Write-Verbose "Connected to SharePoint Online"

            }
            catch {

                Write-Error "An error occurred while attempting to connect to SharePoint Online. Exiting function. Your issue is either with your Internet connection, the SharePoint Online Administration site URL you specified is incorrect, the credentials for the SharePoint Online Administration site you specified are incorrect or the SharePoint Online system itself is experiencing a problem. Please read the error message below carefully for clues on how to resolve your issue."
                Write-Error $_
                Exit

            }
        }

        Write-Verbose "Leaving Process Block: Connect-SPOnline"

    }
}
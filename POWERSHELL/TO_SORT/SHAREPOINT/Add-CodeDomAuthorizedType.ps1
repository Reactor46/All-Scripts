<#
 This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.  
 THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
 INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  
 We grant you a nonexclusive, royalty-free right to use and modify the sample code and to reproduce and distribute the object 
 code form of the Sample Code, provided that you agree: 
    (i)   to not use our name, logo, or trademarks to market your software product in which the sample code is embedded; 
    (ii)  to include a valid copyright notice on your software product in which the sample code is embedded; and 
    (iii) to indemnify, hold harmless, and defend us and our suppliers from and against any claims or lawsuits, including 
          attorneys' fees, that arise or result from the use or distribution of the sample code.
 Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within 
              the Premier Customer Services Description.
 ----------------------------------------------------------
 History
 ----------------------------------------------------------
 10/15/2018 - Added three additional authorized types
 09/18/2018 - Added an update to allow customers using Nintex to use the new IncludeNintexWorkflow switch to automatically add
              the necessary authorizedType required for Nintex
 09/17/2018 - Updated to match "final update" post
 
   REFERENCE:
    
    https://support.microsoft.com/en-us/help/4465015/sharepoint-workflows-stop-after-cve-2018-8421-security-update
    https://blogs.msdn.microsoft.com/rodneyviana/2018/09/13/after-installing-net-security-patches-to-address-cve-2018-8421-sharepoint-workflows-stop-working/
    https://blogs.msdn.microsoft.com/rodneyviana/2018/10/12/step-by-step-video-on-how-to-fix-the-sharepoint-workflow/
  SUMMARY: 
    
    This script leverages the native SharePoint SPWebConfigModification API to deploy new updates to the web.config file for
    each web application on each server in the farm.  Servers added a later date will also get the updates applied because the API 
    configuration is persisted in the config database.  This API does not update the web.config for the central administration web application. 
    If you are running workflows on the central admin web application, you will need to manually update the web.config using the steps in the 
    referenced blog.
==============================================================
#>

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null

function Add-CodeDomAuthorizedType
{
    <#
    .Synopsis
       Adds the necessary authorizedType elements to all web.config files for all non-central admin web applications 
    .DESCRIPTION
       Adds the necessary authorizedType elements to all web.config files for all non-central admin web applications 
    .EXAMPLE
       Add-CodeDomAuthorizedType
    .EXAMPLE
       Add-CodeDomAuthorizedType -IncludeNintexWorkflow
    #>
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$false)][switch]$IncludeNintexWorkflow
    )

    begin
    {
        $updateRequired = $false

        $farmMajorVersion = (Get-SPFarm -Verbose:$false ).BuildVersion.Major
        $contentService = [Microsoft.SharePoint.Administration.SPWebService]::ContentService

        $authorizedTypes = @()
 
        if( $farmMajorVersion -le 14)
        {
            $systemAssemblyVersion = "2.0.0.0"
            $targetParentPath      = "configuration/System.Workflow.ComponentModel.WorkflowCompiler/authorizedTypes"
        }
        else
        {
            $systemAssemblyVersion = "4.0.0.0"
            $targetParentPath      = "configuration/System.Workflow.ComponentModel.WorkflowCompiler/authorizedTypes/targetFx[@version='v4.0']"
        }

        if($IncludeNintexWorkflow.IsPresent)
        {
            $authorizedTypes += New-Object PSCustomObject -Property @{
                Assembly  = "System, Version=$systemAssemblyVersion, Culture=neutral, PublicKeyToken=b77a5c561934e089"
                Namespace = "System.CodeDom"
                TypeName  = "CodeTypeReferenceExpression"
            }
        }
        
        $authorizedTypes += New-Object PSCustomObject -Property @{
            Assembly  = "System, Version=$systemAssemblyVersion, Culture=neutral, PublicKeyToken=b77a5c561934e089"
            Namespace = "System.CodeDom"
            TypeName  = "CodeBinaryOperatorExpression"
        } 

        $authorizedTypes += New-Object PSCustomObject -Property @{
            Assembly  = "System, Version=$systemAssemblyVersion, Culture=neutral, PublicKeyToken=b77a5c561934e089"
            Namespace = "System.CodeDom"
            TypeName  = "CodePrimitiveExpression"
        } 

        $authorizedTypes += New-Object PSCustomObject -Property @{
            Assembly  = "System, Version=$systemAssemblyVersion, Culture=neutral, PublicKeyToken=b77a5c561934e089"
            Namespace = "System.CodeDom"
            TypeName  = "CodeMethodInvokeExpression"
        } 

        $authorizedTypes += New-Object PSCustomObject -Property @{
            Assembly  = "System, Version=$systemAssemblyVersion, Culture=neutral, PublicKeyToken=b77a5c561934e089"
            Namespace = "System.CodeDom"
            TypeName  = "CodeMethodReferenceExpression"
        } 

        $authorizedTypes += New-Object PSCustomObject -Property @{
            Assembly  = "System, Version=$systemAssemblyVersion, Culture=neutral, PublicKeyToken=b77a5c561934e089"
            Namespace = "System.CodeDom"
            TypeName  = "CodeFieldReferenceExpression"
        } 

        $authorizedTypes += New-Object PSCustomObject -Property @{
            Assembly  = "System, Version=$systemAssemblyVersion, Culture=neutral, PublicKeyToken=b77a5c561934e089"
            Namespace = "System.CodeDom"
            TypeName  = "CodeThisReferenceExpression"
        } 

        $authorizedTypes += New-Object PSCustomObject -Property @{
            Assembly  = "System, Version=$systemAssemblyVersion, Culture=neutral, PublicKeyToken=b77a5c561934e089"
            Namespace = "System.CodeDom"
            TypeName  = "CodePropertyReferenceExpression"
        }

        # added 10/15/2018 to match Nov 2018 CU
        $authorizedTypes += New-Object PSCustomObject -Property @{
            Assembly  = "System.Workflow.Activities, Version=$systemAssemblyVersion, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
            Namespace = "System.Workflow.Activities.Rules"
            TypeName  = "RuleDefinitions"
        }

        # added 10/15/2018 to match Nov 2018 CU
        $authorizedTypes += New-Object PSCustomObject -Property @{
            Assembly  = "System.Workflow.Activities, Version=$systemAssemblyVersion, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
            Namespace = "System.Workflow.Activities.Rules"
            TypeName  = "RuleExpressionCondition"
        }

        #  this should exist in web.config already
        #$authorizedTypes += New-Object PSCustomObject -Property @{
        #    Assembly  = "Microsoft.SharePoint.WorkflowActions, Version=$farmMajorVersion.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"
        #    Namespace = "Microsoft.SharePoint.WorkflowActions"
        #    TypeName  = "*"
        #}
    }
    process
    {
        foreach( $authorizedType in $authorizedTypes )
        {
            $netFrameworkConfig = New-Object Microsoft.SharePoint.Administration.SPWebConfigModification
            $netFrameworkConfig.Path     = $targetParentPath
            $netFrameworkConfig.Name     = "authorizedType[@Assembly='$($authorizedType.Assembly)'][@Namespace='$($authorizedType.Namespace)'][@TypeName='$($authorizedType.TypeName)'][@Authorized='True']" 
            $netFrameworkConfig.Owner    = "NetFrameworkAuthorizedTypeUpdate"
            $netFrameworkConfig.Sequence = 0
            $netFrameworkConfig.Type     = [Microsoft.SharePoint.Administration.SPWebConfigModification+SPWebConfigModificationType]::EnsureChildNode
            $netFrameworkConfig.Value    = "<authorizedType Assembly=`"$($authorizedType.Assembly)`" Namespace=`"$($authorizedType.Namespace)`" TypeName=`"$($authorizedType.TypeName)`" Authorized=`"True`"/>"
            
            if( -not ($contentService.WebConfigModifications | ? { $_.Value -eq $netFrameworkConfig.Value }) )
            {
                Write-Verbose "Adding Authorized Type: $($netFrameworkConfig.Value)"

                $contentService.WebConfigModifications.Add($netFrameworkConfig);
                $updateRequired = $true
            }
            else
            {
                Write-Verbose "Authorized Type Exists: $($netFrameworkConfig.Value)"
            }
        }

        if( $updateRequired )
        {
            Write-Verbose "Updating web.configs"
            $contentService.Update()
            $contentService.ApplyWebConfigModifications();
        }
    }
    end
    {
    }    
}

function Remove-CodeDomAuthorizedType
{
    <#
    .Synopsis
       Removes any web configuration entires owned by "NetFrameworkAuthorizedTypeUpdate"
    .DESCRIPTION
       Removes any web configuration entires owned by "NetFrameworkAuthorizedTypeUpdate"
    .EXAMPLE
        Remove-CodeDomAuthorizedType
    #>
    [CmdletBinding()]
    param()

    begin
    {
        $contentService = [Microsoft.SharePoint.Administration.SPWebService]::ContentService
    }
    process
    {
        $webConfigModifications = @($contentService.WebConfigModifications | ? { $_.Owner -eq "NetFrameworkAuthorizedTypeUpdate" })

        foreach ( $webConfigModification in $webConfigModifications ) 
        {
            Write-Verbose "Found instance owned by NetFrameworkAuthorizedTypeUpdate"
            $contentService.WebConfigModifications.Remove( $webConfigModification ) | Out-Null
        }
        
        if( $webConfigModifications.Count -gt 0 )
        {
            $contentService.Update()
            $contentService.ApplyWebConfigModifications()
        }
    }
    end
    {
    }    
}

# will get the timerjob responsible for the web.config change deployment
# Get-SPTimerJob | ? { $_.Name -eq "job-webconfig-modification" }

# adds the updates to the farm, only needs to be run once per farm.
Add-CodeDomAuthorizedType -Verbose

# remove # below if you need to remove the web.config updates, you can with this function to retract the changes
# Remove-CodeDomAuthorizedType -Verbose
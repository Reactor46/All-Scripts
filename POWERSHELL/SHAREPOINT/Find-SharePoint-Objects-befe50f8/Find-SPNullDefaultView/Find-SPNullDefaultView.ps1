function Find-SPNullDefaultView{
<#
	.SYNOPSIS
		Displays all SharePoint Objects without a DefaultView defined.
	
	.DESCRIPTION
		Useful for identifying objects where the Default View has been deleted, or objects created through migration tools 
        that failed to create all aspects of the SPList object.  Returns the following SPList Object attributes:
        Title
        ItemCount
        Author
        LastItemModifiedDate
        ParentWebUrl
        BaseType
        RootFolder
        Hidden

	
	.EXAMPLE
        Find-SPNullDefaultView <url>
        Find-SPNullDefaultView -url "http://site/subsite"
        Find-SPNullDefaultView <url> | Format-Table Title, BaseType, Hidden, ItemCount, Author, LastItemModifiedDate, ParentWebUrl, RootFolder -AutoSize
        Find-SPNullDefaultView <url> | Select-Object Title, BaseType, Hidden, ItemCount, Author, LastItemModifiedDate, ParentWebUrl, RootFolder | Out-GridView
        

    .REQUIREMENTS
        Microsoft.SharePoint Assembly Class
        http://msdn.microsoft.com/en-us/library/microsoft.sharepoint.aspx

    .NOTES
        NAME: Find-SPNullDefaultView
        AUTHOR: Marc Carter
        LASTEDIT: 16-Sep-2013
        KEYWORDS: Windows SharePoint Services 3, MOSS, MOSS 2007, WSS3
        
        SPList.Hidden property: Hidden SPList Objects: http://msdn.microsoft.com/en-us/library/microsoft.sharepoint.splist.hidden.aspx
        Can easily be modified to produce other results by modifying search query 'If($List.DefaultView -eq $null)'

#>
    [cmdletbinding()]	    
    param(
        [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True,Mandatory=$True)]
        [string[]]$url
    )
    Begin{
        $const_verbosepreference = $verbosepreference
        $verbosepreference = "Continue"
        # Attempt to load SharePoint Assembly Class before proceeding
        Try { [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") } 
        Catch { 
            Write-Warning "Failed to Load REQUIRED .NET Framework Assembly Class...exiting script!"
            Write-Warning "http://msdn.microsoft.com/en-us/library/hh537936(v=office.14).aspx"
            Break 
        }
        $array = @()
    }
    Process{
        Try{
            $SiteCollection = New-Object Microsoft.SharePoint.SPSite($url)
            Foreach($Site in $SiteCollection.AllWebs){ 
                Write-Verbose "$($Site.Url)"
                Foreach ($List in $Site.Lists){ 
                    Try{ 
                        If($List.DefaultView -eq $null){ 
                            $props = @{
                                Title=$($List.Title)
                                ItemCount=$($List.ItemCount)
                                Author=$($List.Author)
                                LastItemModifiedDate=$($List.LastItemModifiedDate)
                                ParentWebUrl=$($List.ParentWebUrl)
                                BaseType=$($List.BaseType)
                                RootFolder=$($List.RootFolder) 
                                Hidden=$($List.Hidden) 
                            }
                            $array += New-Object PSObject -property $props
                        } 
                    } Catch { Write-Verbose "Unable to load $List.DefaultViewUrl" } 
                } 
            } 
            $SiteCollection.Dispose(); 
        } Catch { Write-Verbose "Unable to locate site for $($url)" }
    }
    End{
        $array
        $verbosepreference = $const_verbosepreference 
    }
}

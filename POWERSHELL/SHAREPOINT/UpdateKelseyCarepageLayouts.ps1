$url = "https://uat.www.kelseycare.com"   # Site collection
$comment = "Batch PageLayout Update"   # Publishing comment

# Function: Update-SPPagesPageLayout
# Description: Update a single page in a Publishing Web
# Parameters: publishingPage, pageLayout, comment
function Update-SPPagesPageLayout ([Microsoft.SharePoint.Publishing.PublishingPage]$publishingPage,
    [Microsoft.SharePoint.Publishing.PageLayout] $pageLayoutNew, [string]$comment)
{
    Write-Host "Updating the page:" $publishingPage.Name "to Page Layout:" $pageLayoutNew.Title
    $publishingPage.CheckOut();
    $publishingPage.Layout = $pageLayoutNew;
    $publishingPage.ListItem.Update();
    $publishingPage.CheckIn($comment);
    #if ($publishingPage.ListItem.ParentList.EnableModeration)
    #{
    #    $page.ListItem.File.Approve("Publishing Page Layout correction");
    #}
}

# Function: Update-AllSPPagesPageLayouts
# Description: Loop through all the pages in a Publishing Web and update their page layout
# Parameters: web, pageLayoutCurrent, pageLayoutNew, comment
# comment Comment to accompany the checkin
Function Update-AllSPPagesPageLayouts ([Microsoft.SharePoint.SPWeb]$web, [Microsoft.SharePoint.Publishing.PageLayout]$pageLayoutCurrent,
    [Microsoft.SharePoint.Publishing.PageLayout]$pageLayoutNew, [string]$comment)
{
    #Check if this is a publishing web
    if ([Microsoft.SharePoint.Publishing.PublishingWeb]::IsPublishingWeb($web) -eq $true)
    {
      $pubweb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web);
      $pubcollection=$pubweb.GetPublishingPages()
      #Go through all pages checking for pages with the "current" page layout
      for($i=0; $i -lt $pubcollection.count; $i++)
      {
        if($pubcollection[$i].Layout.Title -eq $pageLayoutCurrent.Title)
        {
            Update-SPPagesPageLayout $pubcollection[$i] $pageLayoutNew $comment
        }
      }
    }
    $web.Close();
}

# Check Parameters
#if(($args[0] -ne $null) -and ($args[1] -ne $null))
#{
    #Write-Host "** Update Layout Pages from-" $args[0] "-to-" $args[1] "-on URL" $url
    $pageLayoutNameCurrent = "Subsite Body Page";
    $pageLayoutNameNew = "Kelsey Care Detail";

    $site = new-object Microsoft.SharePoint.Publishing.PublishingSite(Get-SPSite $url)

    Write-Host "Checking if both page layouts exist in the site..."
    # Check if the current pagelayout exists in this site collection
    $pageLayouts = $site.GetPageLayouts($true);

    $pageLayouts | ForEach-Object {
        if ($_.Title -eq $pageLayoutNameCurrent)
        {
            Write-Host "Found CURRENT page layout: " $pageLayoutNameCurrent
            $pageLayoutCurrent = $_;
        }
    }

    # Check if the new pagelayout exists in this site collection
    $pageLayouts | ForEach-Object {
        if ($_.Title -eq $pageLayoutNameNew)
        {
            Write-Host "Found NEW page layout: " $pageLayoutNameNew
            $pageLayoutNew = $_;
        }
    }      

    # Do not continue if the either pageLayout does not exist
    if(($pageLayoutCurrent -ne $null) -and ($pageLayoutNew -ne $null))
    {
        # Update all subsites
        #if($args[2] -eq "-all")
        #{
         $site.Site.allwebs | foreach {
            Write-Host "Checking Web: " $_.Title
            Update-AllSPPagesPageLayouts $_ $pageLayoutCurrent $pageLayoutNew $comment
            }
        #}
        #else
        #{
         #$site.rootweb | foreach {
         #   Write-Host "Checking Web: " $_.Title
         #   Update-AllSPPagesPageLayouts $_ $pageLayoutCurrent $pageLayoutNew $comment
         #   }
        #}
    }
    Write-Host "**Done"
#}
#else
#{
#    Write-Host "Missing arguments.  Please check your parameters"
#}
#End
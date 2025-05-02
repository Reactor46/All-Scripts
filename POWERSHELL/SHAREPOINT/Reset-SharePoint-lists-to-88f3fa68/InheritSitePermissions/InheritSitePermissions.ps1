#========================================================================
# Code File: InheritSitePermissions.ps1
# Created On: 1/5/2012
# Created By: Cornelius J. van Dyk
#========================================================================

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint");
$url = $args[0];
$lib = $args[1];
$all = $args[2];
$site = new-object microsoft.sharepoint.spsite($url);
write-host "Resetting list permissions for " $lib " to inherit from parent.";
for ($s=0;$s -lt $site.allwebs.count;$s++)
{
  write-host "Checking " $site.allwebs[$s].url " for target list...";
  for ($l=0;$l -lt $site.allwebs[$s].lists.count;$l++)
  {
    if ($site.allwebs[$s].lists[$l].title -eq $lib)
    {
      write-host "Found the " $site.allwebs[$s].lists[$l].title " target list!";
      $site.allwebs[$s].lists[$l].resetroleinheritance();
      if ($all -ne "-all")
      {
        exit;
      }
    }
  }
}
$site.dispose();
write-host "Done."

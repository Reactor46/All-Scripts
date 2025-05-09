$siteurl=""
$listurl = ""
$fldName = ""

$session = Get-SPTaxonomySession -Site $siteurl
$defaultKeywordStore  = $session.DefaultKeywordsTermStore
$SystemGroup = $defaultKeywordStore.Groups["System"]
$KeywordsTermSet = $SystemGroup.TermSets["Keywords"]
$existedKeywords = $KeywordsTermSet.GetAllTerms() |Select Name

$list = Get-SPList $listurl 
$list.Items | % {
$k =  $_[$fldName];
write-output $k
IF ($existedKeywords | ?{$_.Name -eq $k}) {"existed"} ELSE {$KeywordsTermSet.CreateTerm($k,1033)}
write-output "_________"
}

$defaultKeywordStore.CommitAll()


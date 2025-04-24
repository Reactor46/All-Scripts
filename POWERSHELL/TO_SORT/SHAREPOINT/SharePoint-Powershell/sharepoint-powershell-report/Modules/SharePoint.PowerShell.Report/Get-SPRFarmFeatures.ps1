function Get-SPRFarmFeatures
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPFarmFeatures,
    [Parameter(Mandatory = $true)]
    [object[]]$SPWebServicesAdministration
  )
	
$FarmFeatures=$SPFarmFeatures | ? {$_.Scope.value -eq "Farm"}
$ActiveFarmFeatures=$SPWebServicesAdministration

  foreach ($FarmFeature in $FarmFeatures)
  {
    $isactive=$false
    foreach ($ActiveFarmFeature in $ActiveFarmFeatures) {
   
    if ($FarmFeature.id -eq $ActiveFarmFeature.definitionid) {$isactive=$true} 
    
    }
    if ($FarmFeature.SolutionID -ne '7ed6cd55-b479-4eb7-a529-e99a24c10bd3' -and $FarmFeature.SolutionID  -ne '00000000-0000-0000-0000-000000000000'-and $FarmFeature.SolutionID  -ne 'a992b0f0-7b6b-4315-b687-649dcaeef726') {$iscustom=$true} 
    else {$iscustom=$false}
    $properties = [ordered]@{
      'Title'  = $FarmFeature.DisplayNameEN
      'Id'     = $FarmFeature.id
      'SolutionID' = $FarmFeature.SolutionID
      'SolutionName'=$xxx
      'Version' = $FarmFeature.Version
      'Active' = $isactive
      'Hidden' = $FarmFeature.Hidden
      'Custom' = $iscustom
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output

    
  }
}
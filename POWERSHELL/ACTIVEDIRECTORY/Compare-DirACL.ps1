<# 
.SYNOPSIS 
Compares the Access Control Lists for two directories. 
 
.DESCRIPTION 
Function to comapre two directories' access permissions. 
 
.PARAMETER ReferenceDir 
The reference directory to be compared to.

.PARAMETER DifferenceDir
The directory to compare to the reference directory. 
 
.EXAMPLE 
Format output as a list (default behavior)
PS> .\Compare-DirACL.ps1 -ReferenceDir C:\Temp\DirOne -DifferenceDir C:\Temp\DirTwo

.EXAMPLE
Format output as a table
PS> .\Compare-DirACL.ps1 -ReferenceDir C:\Temp\DirOne -DifferenceDir C:\Temp\DirTwo | ft
#> 
 
param 
( 
    [parameter()][string] $ReferenceDir,
    [parameter()][string] $DifferenceDir 
) 

$ReferenceACL = Get-Acl $ReferenceDir
$DifferenceACL = Get-Acl $DifferenceDir
$Comparison = Compare-Object -ReferenceObject $ReferenceACL -DifferenceObject $DifferenceACL -Property access | Select sideindicator -ExpandProperty access
return $Comparison
#Defining variables
$x = 3
$y = 5
$z = 2

#Using logical operators to print out statements describing the above variables
if(($x -and $y -ge 5) -or ($z -eq 2)){
Write-Host 'My numbers are high enough'
}
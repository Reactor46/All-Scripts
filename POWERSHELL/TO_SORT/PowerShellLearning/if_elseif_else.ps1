#Creating the variable and assigning a value of 57 to it
$var = 57

#Crafting an if, elseif, else statement to print out an appropriate message stating whether it is less than, equal to or greater than 57.
if($var -lt 57){
Write-Host 'var is less than 57'
} elseif($var -eq 57) {
Write-Host 'var is equal to 57'
} else {
Write-Host 'var is greater than 57'
}
#Create an array of numbers
$numbers = 1,2,3,4,5,6

#Create a variable for the array to act upon
$starting_number = 2

#Make it so that each number in the numbers array is multiplied by the 
foreach ($number in $numbers){
$starting_number *= $number
}

#Print the result to the console
Write-Host 'The total is ' $starting_number
#Creating an array that my while loop will then be able to sift through.
$Beast_Wars_Characters =@('Optimus Primal', 'Dinobot', 'Rhinox')

#Setting $searching to $true to establish that the array is meant to be searched through.
$searching = $true
#While loop is then written to say 'as long as the program is searching, print these messages depending on what's in the loop'
while($searching = $true){
if($Beast_Wars_Characters -contains 'Megatron'){
Write-Host 'You have Megtaron on your side'
break
}
else{
Write-Host 'You have not got Megatron.'
break
}
} 
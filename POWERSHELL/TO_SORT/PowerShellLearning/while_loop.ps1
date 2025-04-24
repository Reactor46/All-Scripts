#Creating an array to establish what the while loop can search through.
$player =@('Rattrap', 'Dinobot', 'Rhinox')

#Establishing that searching is set to true in order to look for the value Megatron
$searching = $true
#The while loop follows, which will print different statements depending on whether or not Megatron appears in the array
while ($searching) {
if ($player -contains 'Megatron' -or 'Rattrap'){
    Write-Host 'The player has Megatron or Rattrap in their team'
    $searching = $false
}
else {
    Write-Host 'The player has not got Megatron or Rattrap in their team. Searching again...'
}
}
#Array to store the names of predacons
$predacons=@('Megatron', 'Scorponok', 'Waspinator')

#Counter variable to enable me to cycle through the array
$counter = 0

#do until loop which will write to the console that each of the above transformers is a predacon. 
do {

Write-Host $predacons[$counter] ' is a predacon'
$counter++

} until ($counter -eq 3)
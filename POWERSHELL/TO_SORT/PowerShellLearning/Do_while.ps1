#Creating an array of 5 evil beast wars transformers
$predacons =@('Megatron', 'Scorponok', 'Tarantulas', 'Waspinator', 'Terrorsaur')

#Defining a counter variable. This will enable me to cycle through the predacons.
$counter = 0

#do while loop which will write to the console that each of the above transformers is a predacon. 

Do{
Write-Host $predacons[$counter] 'is a predacon!'
$counter++
}
#As there are 5, it will print these names until it gets to the fifth one.
While($counter -ne 5)

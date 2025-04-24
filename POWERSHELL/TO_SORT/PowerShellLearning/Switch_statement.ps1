#Establishing the Beast wars character variable as Megatron.
$Beast_Wars_Character = 'Megatron'


#Given that 'Megatron' is the value in the above variable, this switch statement will print 'Megatron' to the console.
#If I changed the above to 'Terrorsaur', it would print Terrorsaur
switch($Beast_Wars_Character)
{
"Megatron" {Write-Host 'Megatron'}
"Terrorsaur"{Write-Host 'Terrorsaur'}
"Scorponok" {Write-Host 'Scorponok'}
"Tarantulas" {Write-Host 'Tarantulas'}
"Waspinator" {Write-Host 'Waspinator'}
default {Write-Host 'I have no idea who this character is'}
}
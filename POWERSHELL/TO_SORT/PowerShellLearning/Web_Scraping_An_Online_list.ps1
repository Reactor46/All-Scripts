#Defining a variable to store this list

$pokemon = Invoke-WebRequest -Uri https://bulbapedia.bulbagarden.net/wiki/List_of_Pok%C3%A9mon_by_National_Pok%C3%A9dex_number#List_of_Pok.C3.A9mon_by_National_Pok.C3.A9dex_number -UseBasicParsing

$pokemon.Links | Select-Object href | Out-File Pokemonsters.txt

Get-Content Pokemonsters.txt

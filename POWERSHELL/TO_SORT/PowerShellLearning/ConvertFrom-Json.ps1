#Storing the json in a variable
$json = Get-Content -Path 'C:\Users\silad\Documents\BeastWarsCharacters.json' -Raw
#Creating a new variable to hold the contents of rhe json file after they've been converting into PowerShell objects
$data = ConvertFrom-Json -InputObject $json

#Present the contents on screen, following their conversion
$data 


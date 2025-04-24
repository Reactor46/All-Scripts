# Storing the json in a variable
$json = Get-Content -Path 'C:\Users\silad\Documents\BeastWarsCharacters.json' -Raw

# Creating a new variable to hold the contents of the json file after they've been converted into PowerShell objects
$data = ConvertFrom-Json -InputObject $json

# Display only the names of the characters using dot notation
foreach ($character in $data) {
    $character.Name
}

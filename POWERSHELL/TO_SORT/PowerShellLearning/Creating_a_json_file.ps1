# Define the file path
$jsonFilePath = "C:\Users\silad\Documents\BeastWarsCharacters.json"

# Create an array of Beast Wars characters with their allegiances
$beastWarsCharacters = @(
    @{
        Name = "Optimus Primal"
        Allegiance = "Maximal"
    },
    @{
        Name = "Megatron"
        Allegiance = "Predacon"
    },
    @{
        Name = "Cheetor"
        Allegiance = "Maximal"
    },
    @{
        Name = "Rattrap"
        Allegiance = "Maximal"
    },
    @{
        Name = "Tarantulas"
        Allegiance = "Predacon"
    }
)

# Convert the array to JSON and save it to a file
$beastWarsCharacters | ConvertTo-Json | Set-Content -Path $jsonFilePath

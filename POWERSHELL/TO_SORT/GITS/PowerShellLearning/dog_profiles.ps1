# Profiles for my favourite dogs
# Define dog profiles using hashtables
$dog1 =@{
Name = "Brea"
Age = "6"
Breed = "German Shepherd"
Personality = "Playful"
Signature_Move = "Twirling"
}

$dog2 =@{
Name = "Kane"
Age = "9"
Breed = "German Shepherd"
Personality = "Talkative"
Signature_Move = "Rolling onto his back and kicking his legs up"
}

$dog3 =@{
Name = "Nook"
Age = "10"
Breed = "Siberian Husky"
Personality = "Playful"
Signature_Move = "Running away from everyone"
}


# Function to display the profiles
function Get-DogProfile {
    param (
        [hashtable]$Profile
    )

    
    Write-Host "Name: $($Profile.Name)"
    Write-Host "Age: $($Profile.Age)"
    Write-Host "Breed: $($Profile.Breed)"
    Write-Host "Personality: $($Profile.Personality)"
    Write-Host "Signature Move: $($Profile.Signature_Move)"
    Write-Host ""
}

# Call the function for each dog
Get-DogProfile $dog1
Get-DogProfile $dog2
Get-DogProfile $dog3

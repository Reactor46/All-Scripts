# Mad Lib Game

# Prompt the user to enter words for the Mad Lib
$adjective = Read-Host "Enter an adjective:"
$adjective_2 = Read-Host "Enter an adjective:"
$adjective_3 = Read-Host "Enter an adjective:"
$noun = Read-Host "Enter a noun:"
$verb = Read-Host "Enter a verb:"
$adverb = Read-Host "Enter an adverb:"
$place = Read-Host "Enter a place:"
$animal = Read-Host "Enter an animal:"

# Create the Mad Lib story
$madLibStory = @"
Once upon a time, there was a $adjective, $adjective_2 and $adjective_3 $noun who loved to $verb $adverb.
One day, they decided to go to $place with their pet $animal.
"@

# Display the Mad Lib story
Write-Host $madLibStory

# Wait for user input to prevent the console from closing immediately
Read-Host "Press Enter to exit"

# Set the welcome message
$welcomeMessage = "Welcome to the Shayan Awards! Are you ready to get started?"

# Display a message box with the welcome message
[System.Windows.Forms.MessageBox]::Show($welcomeMessage, "The Shayan Awards", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

# Wait for user to press Enter to continue
Write-Host "Press Enter to continue..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").KeyDownCharacter

# First Award Category
$awardMessage1 = "The first one is the Clever Clogs award. The award goes to....."
# Display the award message
Write-Host $awardMessage1
# Wait for user to press Enter to reveal the winner
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").KeyDownCharacter
# Announce the winner
$winner1 = "Damir Derd"
Write-Host "Winner: $winner1"
# Honorable Mentions
$honorableMentions1 = "Honorable Mentions: James Gilmore, Brad Phipps, Roy Bentley"
Write-Host $honorableMentions1

# Second Award Category
$awardMessage2 = "The second category is the 'Cup of Kindness award'. The award goes to....."
# Display the award message
Write-Host $awardMessage2
# Wait for user to press Enter to reveal the winner
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").KeyDownCharacter
# Announce the winner
$winner2 = "Catherine King"
Write-Host "Winner: $winner2"
# Honorable Mentions
$honorableMentions2 = "Honorable Mentions: Reggie Mitchell, Ross Thompson, Severine Vidal"
Write-Host $honorableMentions2

# Third Award Category
$awardMessage3 = "The third category is the 'Humour Homunculus award'. The award goes to....."
# Display the award message
Write-Host $awardMessage3
# Wait for user to press Enter to reveal the winner
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").KeyDownCharacter
# Announce the winner
$winner3 = "Amelia Kent"
Write-Host "Winner: $winner3"
# Honorable Mentions
$honorableMentions3 = "Honorable Mentions: Alejandro Lemus-Gomez, Marcin Szymor, Paolina White"
Write-Host $honorableMentions3

# Fourth Award Category
$awardMessage4 = "The 'Steady Eddie' award goes to....."
# Display the award message
Write-Host $awardMessage4
# Wait for user to press Enter to reveal the winner
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").KeyDownCharacter
# Announce the winner
$winner4 = "Dan Cochrane"
Write-Host "Winner: $winner4"
# Honorable Mentions
$honorableMentions4 = "Honorable Mentions: Katerina Vlckova, Vignesh Umapathy, Chris Saner"
Write-Host $honorableMentions4

# Pause before closing the window
Write-Host "Press Enter to close the Shayan Awards..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").KeyDownCharacter

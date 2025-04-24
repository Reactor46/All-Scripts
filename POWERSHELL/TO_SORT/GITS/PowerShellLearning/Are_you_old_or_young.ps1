# Prompt the user for their opinion
$opinion = Read-Host "Do you think Eddie Murphy is funny? (yes/no)"

# Check the user's response and display a message accordingly
if ($opinion -eq "yes") {
    $message = "You are over 35 (old as shit)."
} elseif ($opinion -eq "no") {
    $message = "You are under 35."
} else {
    $message = "Invalid response. Please enter 'yes' or 'no'."
}

# Display the message
Write-Host $message

# Wait for user input to prevent the console from closing immediately
Read-Host "Press Enter to exit"

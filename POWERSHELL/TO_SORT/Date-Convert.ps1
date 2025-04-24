# Define the ISO 8601 date string
$isoDate = "2026-04-12T14:40:02Z"

# Convert to DateTime object
$dateTime = [DateTime]::Parse($isoDate)

# Format the DateTime object to a human-readable string
$humanReadableDate = $dateTime.ToString("dddd, MMMM dd, yyyy hh:mm tt")

# Output the human-readable date
$humanReadableDate
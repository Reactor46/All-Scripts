# Establish the directory that I want to remove all files over 30 days old from
$directory = 'C:\Users\User\Pictures\Video_Projects'

# Get the current date
$currentDate = Get-Date

# Calculate the date 30 days ago dynamically
$thresholdDate = $currentDate.AddDays(-30)

# Pull up files in the directory and subdirectories older than 30 days based on LastWriteTime
$oldFiles = Get-ChildItem -Path $directory -File -Recurse | Where-Object { $_.LastWriteTime -lt $thresholdDate }

# Check if there are any files to delete
if ($oldFiles.Count -gt 0) {
    foreach ($file in $oldFiles) {
        # Remove the file
        Remove-Item -Path $file.FullName -Force
        Write-Output "Removed file: $($file.FullName)"
    }
    Write-Output "Files older than 30 days have been removed."
} else {
    Write-Output "No files older than 30 days found in the specified directory."
}

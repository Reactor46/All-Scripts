# Define the path to search
$searchPath = "C:\"

# Search for jackson-databind JAR files
$files = Get-ChildItem -Path $searchPath -Filter "jackson-databind*.jar" -Recurse -ErrorAction SilentlyContinue

# Check if any files were found
if ($files.Count -eq 0) {
    Write-Output "No jackson-databind files found in the specified path."
} else {
    # Iterate over each found file and report its version
    foreach ($file in $files) {
        # Extract the version from the filename
        if ($file.Name -match "jackson-databind-(\d+\.\d+\.\d+).jar") {
            $version = $matches[1]
            Write-Output "File: $($file.FullName)"
            Write-Output "Version: $version"
        } else {
            # If version is not in the filename, check inside the JAR
            Write-Output "File: $($file.FullName)"
            Write-Output "Version information not found in the filename."
        }
    }
}
# Define the path to search
$searchPath = "C:\"

# Function to get version from filename
function Get-VersionFromFilename {
    param ($filename)
    if ($filename -match "jackson-databind-(\d+\.\d+\.\d+).jar") {
        return $matches[1]
    } else {
        return "Version information not found in the filename."
    }
}

# Search for jackson-databind JAR files
$jarFiles = Get-ChildItem -Path $searchPath -Filter "jackson-databind*.jar" -Recurse -ErrorAction SilentlyContinue

# Search for ZIP files and inspect their contents
$zipFiles = Get-ChildItem -Path $searchPath -Filter "*.zip" -Recurse -File -ErrorAction SilentlyContinue

# Check if any JAR files were found directly
if ($jarFiles.Count -gt 0) {
    foreach ($file in $jarFiles) {
        $version = Get-VersionFromFilename -filename $file.Name
        Write-Output "File: $($file.FullName)"
        Write-Output "Version: $version"
    }
}

# Check if any ZIP files were found
if ($zipFiles.Count -gt 0) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

    foreach ($zipFile in $zipFiles) {
        $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($zipFile.FullName)
        $zipArchive.Entries | ForEach-Object {
            if ($_.FullName -match "jackson-databind-(\d+\.\d+\.\d+).jar") {
                $version = $matches[1]
                Write-Output "Found in ZIP: $($zipFile.FullName)"
                Write-Output "JAR File: $($_.FullName)"
                Write-Output "Version: $version"
            }
        }
        $zipArchive.Dispose()
    }
}

# If no JAR or ZIP files containing jackson-databind were found
if ($jarFiles.Count -eq 0 -and $zipFiles.Count -eq 0) {
    Write-Output "No jackson-databind files found in the specified path."
}

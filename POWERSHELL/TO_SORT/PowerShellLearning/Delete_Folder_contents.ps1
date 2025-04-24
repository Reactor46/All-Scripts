#This script will delete the contents of my Music folder

#Storing the path to my Music folder in a variable
$folder = "C:\Users\silad\Music"

#Check if the folder exists before trying to remove its contents
if (Test-Path $folder -PathType Container){

try {
#Get all items and files in the music folder
$items = Get-ChildItem $folder -Force

#For each loop to cycle through the contents and delete them
foreach($item in $items){
Remove-Item -Path $item.FullName -Force -Recurse
}

Write-Host "Contents of '$folder' have been deleted."
    } catch {
        Write-Host "Error: $_"
    }
} else {
    Write-Host "Folder '$folder' not found."
}


$Backgrounds = ("\\lasfs03\software\Backgrounds")
$CurrentWallpaperPath = Get-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name Wallpaper
$WallpaperFileName = "wallpaper1680x1050.jpg"
$DestinationFolderPath = "$Env:USERPROFILE"

function Update-Wallpaper {

    Remove-Item -Path "$($env:APPDATA)\Microsoft\Windows\Themes\CachedFiles" -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$($env:APPDATA)\Microsoft\Windows\Themes\*" -Force -ErrorAction SilentlyContinue

    if ((Test-Path "$DestinationFolderPath\$WallpaperFileName") -eq $False) {
        Copy-Item -Path ("$Backgrounds\$WallpaperFileName") -Destination "$DestinationFolderPath\$WallpaperFileName" -Force
    } else {
        $CurrentWallpaperHash = Get-FileHash $CurrentWallpaperPath.WallPaper
        $BackgroundsWallpaperHash = Get-FileHash $Backgrounds\$WallpaperFileName
        if ($CurrentWallpaperHash.Hash -ne $BackgroundsWallpaperHash.Hash) {
            Copy-Item -Path ("$Backgrounds\$WallpaperFileName") -Destination "$DestinationFolderPath\$WallpaperFileName" -Force
        }
    }

    Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name Wallpaper -Value "$DestinationFolderPath\$WallpaperFileName"
    Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name TileWallpaper -Value "0"
    Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name WallpaperStyle -Value "2" -Force
 }

 Update-Wallpaper
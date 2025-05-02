# Delete all files created more than 2 days ago.
Remove-FilesCreatedBeforeDate -Path "C:\Some\Directory" -DateTime ((Get-Date).AddDays(-2)) -DeletePathIfEmpty 
# Delete all files that have not been updated in 8 hours.
Remove-FilesNotModifiedAfterDate -Path "C:\Another\Directory" -DateTime ((Get-Date).AddHours(-8))
# Delete a single file if it is more than 30 minutes old.
Remove-FilesCreatedBeforeDate -Path "C:\Another\Directory\SomeFile.txt" -DateTime ((Get-Date).AddMinutes(-30))
# Delete all empty directories in the Temp folder, as well as the Temp folder itself if it is empty.
Remove-EmptyDirectories -Path "C:\SomePath\Temp" -DeletePathIfEmpty
# Delete all empty directories created after Jan 1, 2014 3PM.
Remove-EmptyDirectories -Path "C:\SomePath\WithEmpty\Directories" -OnlyDeleteDirectoriesCreatedBeforeDate ([DateTime]::Parse("Jan 1, 2014 15:00:00"))
# See what files and directories would be deleted if we ran the command.
Remove-FilesCreatedBeforeDate -Path "C:\SomePath\Temp" -DateTime (Get-Date) -DeletePathIfEmpty -WhatIf
# Delete all files and directories in the Temp folder, as well as the Temp folder itself if it is empty, and output all paths that were deleted.
Remove-FilesCreatedBeforeDate -Path "C:\SomePath\Temp" -DateTime (Get-Date) -DeletePathIfEmpty -OutputDeletedPaths
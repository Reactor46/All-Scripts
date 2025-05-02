$file_size = Get-ChildItem -Path "C:\Users\TestFile1.txt" -Recurse -Force | Measure-Object -Property Length -Sum

Write-Host 'The file size is ' $file_size
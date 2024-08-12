$snippetsPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::MyDocuments) + "\WindowsPowerShell\snippets"
if(-not(Test-Path -path $snippetsPath))
{
     New-Item -Path $snippetsPath -itemtype directory | Out-Null
     "Created $snippetsPath"
}
$ewsDirectory = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::MyDocuments) + "\WindowsPowerShell\snippets\Exchange Web Services"
Copy-Item -Path (Get-Location).Path.ToString() -Destination $ewsDirectory -Recurse 
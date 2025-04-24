#In this script, an array is defined to account for different ransomware file extensions. 
$ransomewareExtensions =@(".crypt", ".ecc", ".crypto", ".enc", ".locked")
#Following that, a variable called '$files' is created and the value assigned to it is the contents of the C drive. 
$files = Get-ChildItem -Path "C:\Users" -Recurse -File
#Then, there is a variable called $ransomwareFiles' which contains files within the C drive that have the ransomware extensions. 
$ransomewareFiles = $files | Where-Object {$ransomewareExtensions -ccontains $_.Extension}
#Finally, a statement is printed out saying how many ransomware files were detected.
Write-Host "Detected $(($ransomewareFiles).Count) ransomware files:"
$ransomewareFiles | Select-Object -ExpandProperty FullName 
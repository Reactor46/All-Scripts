Function Get-GitCurrentRelease {
[cmdletbinding()]
Param(
[ValidateNotNullorEmpty()]
[string]$Uri = "https://api.github.com/repos/git-for-windows/git/releases/latest"
)
 
Begin {
    Write-Verbose "[BEGIN  ] Starting: $($MyInvocation.Mycommand)"  
 
} #begin
 
Process {
    Write-Verbose "[PROCESS] Getting current release information from $uri"
    $data = Invoke-Restmethod -uri $uri -Method Get
 
    
    if ($data.tag_name) {
    [pscustomobject]@{
        Name = $data.name
        Version = $data.tag_name
        Released = $($data.published_at -as [datetime])
      }
   } 
} #process
 
End {
    Write-Verbose "[END    ] Ending: $($MyInvocation.Mycommand)"
} #end
 
}

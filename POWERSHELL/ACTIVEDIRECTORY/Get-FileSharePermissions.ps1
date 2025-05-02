dir -Recurse \\fs1\Shared | where { $_.PsIsContainer } |
    % { $path1 = $_.fullname; Get-Acl $_.Fullname |
         % { $_.access | where { $_.IdentityReference -like "ENTERPRISE\J.Carter" } |
             Add-Member -MemberType NoteProperty '.\Application Data' -Value $path1 -passthru }} | 
                Export-Csv "C:\LazyWinAdmin\AccountPermissions.csv"
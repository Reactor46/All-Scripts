$subfolders = get-childitem -attributes D -Recurse
#$subfolders | % { write-host $_.fullname; (get-acl -path $_.fullname).access }
$users = @()

foreach ( $folder in $subfolders )
    {
    $fullpath = $folder.fullname
    Write-Host $fullpath -ForegroundColor Green
    $permissions = (get-acl -path $fullpath).access
    foreach ( $permission in $permissions )
        {
        $user = $permission.identityreference.value
        write-host $user
        if ( $users -contains $user )`
            {
            }
        else
            {
            $users = $users + $user
            }
        }     

    }
Write-Host "All Users: 
$users"
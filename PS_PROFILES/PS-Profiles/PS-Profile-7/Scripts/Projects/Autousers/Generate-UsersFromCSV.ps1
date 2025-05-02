$users = get-content csv

foreach ( $user in $user )
    {
    Generate-User -Firstn $user
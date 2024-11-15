function Get-ByOwner
{
    Get-ChildItem -Recurse C:\ | Get-ACL | Where {$_.Owner -Match $args[0] }
}
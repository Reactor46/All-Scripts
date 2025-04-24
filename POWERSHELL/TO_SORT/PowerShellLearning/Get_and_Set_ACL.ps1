# Listing the file permissions for this txt file
$acl = Get-Acl testfile1.txt

# Define the permission details
$permission = 'desktop-du1jkj8\silad', 'FullControl', 'Allow'

# Create a new access rule
$accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission

# Add the access rule to the ACL
$acl.AddAccessRule($accessRule)

# Set the modified ACL back to the file
Set-Acl testfile1.txt $acl

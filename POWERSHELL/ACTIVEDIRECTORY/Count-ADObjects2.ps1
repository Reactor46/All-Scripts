#Modify domain connection 
$Dom = ‘LDAP://DC=Contoso;DC=corp’ 
$Root = New-Object DirectoryServices.DirectoryEntry $Dom

# Create a selector and start searching from the Root of AD 
$selector = New-Object DirectoryServices.DirectorySearcher 
$selector.SearchRoot = $root

#Count number of Computers 
$adobjcomp= $selector.findall() |` 
where {$_.properties.objectcategory -match "CN=Computer"} 
"There are $($adobjcomp.count) computers in the $($root.name) domain"

#Count number of Users 
$adobjUsers= $selector.findall() |` 
where {$_.properties.objectcategory -match "CN=Person"} 
"There are $($adobjUsers.count) people in the $($root.name) domain"

#Count number of OU’s 
$adobjOU= $selector.findall() |` 
where {$_.properties.objectcategory -match "CN=Organizational-Unit"} 
"There are $($adobjOU.count) OU’s in the $($root.name) domain"

#Count number of Groups 
$adobjgroups= $selector.findall() |` 
where {$_.properties.objectcategory -match "CN=Group"} 
"There are $($adobjgroups.count) groups in the $($root.name) domain"
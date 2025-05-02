


$InputBackupName = Read-Host -Prompt "Write the Name of the Backup"


$Backup = Import-CSV C:\Backup\+$InputBackupName



foreach ($acl in $Backup ) {

   if( $acl.IdentityReference.StartsWith("CN=") -Or $acl.IdentityReference.StartsWith("OU=") -Or $acl.IdentityReference.StartsWith("DC=") ){
        $InputDN = $acl.IdentityReference -replace '/',','

   } else {
        $path = "AD:\$InputDN"

        
        $activedirectoryrights = $acl.ActiveDirectoryRights -replace ';',','
        #Write-Host $activedirectoryrights

        $array = $activedirectoryrights -split ", "


        foreach($test in $array){

            $ace = New-Object System.DirectoryServices.ActiveDirectoryAccessRule $acl.IdentityReference, $test, $acl.AccessControlType
            $aclToApply.AddAccessRule($ace)

        }
        #$ace = New-Object System.DirectoryServices.ActiveDirectoryAccessRule $acl.IdentityReference, $activedirectoryrights, $acl.AccessControlType
        #s$aclToApply.AddAccessRule($ace)
        Set-Acl -Path $path -AclObject $aclToApply
        write-host "ACL Applied"
   }


}

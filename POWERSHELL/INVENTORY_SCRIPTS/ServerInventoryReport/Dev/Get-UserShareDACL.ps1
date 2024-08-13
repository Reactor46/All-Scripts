Function Get-UserShareDACL {
    [cmdletbinding()]
    Param(
        [Parameter()]
        $Computername = $Computername                     
    )                   
    Try {    
        Write-Verbose "Computer: $($Computername)"
        #Retrieve share information from comptuer
        $Shares = Get-CimInstance -ClassName Win32_LogicalShareSecuritySetting -ComputerName $Computername -ErrorAction Stop
        ForEach ($Share in $Shares) {
            $MoreShare = $Share.GetRelated('Win32_Share')
            Write-Verbose "Share: $($Share.name)"
            #Try to get the security descriptor
            $SecurityDescriptor = $Share.GetSecurityDescriptor()
            #Iterate through each descriptor
            ForEach ($DACL in $SecurityDescriptor.Descriptor.DACL) {
                [pscustomobject] @{
                    Computername = $Computername
                    Name = $Share.Name
                    Path = $MoreShare.Path
                    Type = $ShareType[[int]$MoreShare.Type]
                    Description = $MoreShare.Description
                    DACLName = $DACL.Trustee.Name
                    AccessRight = $AccessMask[[int]$DACL.AccessMask]
                    AccessType = $AceType[[int]$DACL.AceType]                    
                }
            }
        }
    }
    #Catch any errors                
    Catch {}                                                    
}

#Set-SQLAlias -AliasName "SP13TESTSQL" -DBServer "DBSQL_SP2013_TEST" -PortNumber 45001
Function Set-SQLAlias
{
param
(
    $AliasName,
    $DBServer, 
    $PortNumber, 
    $RemoteComputer
    #[Switch]$Force #TODO: Implement creating alias in case it doesn't exist
)
    $aliasKeyText = "DBMSSOCN,"+$DBServer+","+$portNumber
 
    ## The Alias is stored in the registry here i'm setting the 
    ## REG path for the Alias 32 & 64
    $Bit32 = "SOFTWARE\Microsoft\MSSQLServer\Client\ConnectTo"
    $bit64 = "SOFTWARE\Wow6432Node\Microsoft\MSSQLServer\Client\ConnectTo"
    
    if ($RemoteComputer)
    {
        $HKLM = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $RemoteComputer)
    }
    else
    {
        $HKLM = get-item HKLM:\
    }

    ## check if the ConnectTo SubFolder 
    ## exist in the registry
    if ($HKLM.OpenSubKey($Bit32) -eq $null -or $HKLM.OpenSubKey($Bit32).GetValue($AliasName) -eq $null)
    {
        <#
        if ($Force)
        {
            $key = $HKLM.OpenSubKey("SOFTWARE\Microsoft\MSSQLServer\Client", $true).CreateSubKey("ConnectTo")
        }
        else 
        {
            
        }
        #>
        Write-Host "32 bit alias $AliasName not found, nothing to update"
    }
    else
    {
        $HKLM.OpenSubKey($Bit32, $true).SetValue($AliasName, $aliasKeyText)
    }

    if ($HKLM.OpenSubKey($Bit64) -eq $null -or $HKLM.OpenSubKey($Bit64).GetValue($AliasName) -eq $null)
    {
        <#
        if ($Force)
        {
            $key = $HKLM.OpenSubKey("SOFTWARE\Wow6432Node\Microsoft\MSSQLServer\Client", $true).CreateSubKey("ConnectTo")
        }
        else 
        {
            
        }
        #>
        Write-Host "64 bit alias $AliasName not found, nothing to update"
    }
    else
    {
        $HKLM.OpenSubKey($Bit64, $true).SetValue($AliasName, $aliasKeyText)
    }
}


Export-ModuleMember Set-SQLAlias
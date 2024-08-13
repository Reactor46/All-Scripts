# Region Helper Functions
function Out-DataTable {
    [CmdletBinding()]
    param([Parameter(Position=0, Mandatory=$true, ValueFromPipeline = $true)] [PSObject[]]$InputObject)

    Begin
    {
    function Get-Type {
        param($type)

        $types = @(
        'System.Boolean',
        'System.Byte[]',
        'System.Byte',
        'System.Char',
        'System.Datetime',
        'System.Decimal',
        'System.Double',
        'System.Guid',
        'System.Int16',
        'System.Int32',
        'System.Int64',
        'System.Single',
        'System.UInt16',
        'System.UInt32',
        'System.UInt64')

        if ( $types -contains $type ) {
            Write-Output "$type"
        }
        else {
            Write-Output 'System.String'
        
        }
    } # Get-Type
        $dt = new-object Data.datatable  
        $First = $true 
    }
    Process
    {
        foreach ($object in $InputObject)
        {
            $DR = $DT.NewRow()  
            foreach($property in $object.PsObject.get_properties())
            {  
                if ($first)
                {  
                    $Col =  New-Object Data.DataColumn  
                    $Col.ColumnName = $property.Name.ToString()  
                    if ($property.value)
                    {
                        if ($property.value -isnot [System.DBNull]) {
                            $Col.DataType = [System.Type]::GetType("$(Get-Type $property.TypeNameOfValue)")
                         }
                    }
                    $DT.Columns.Add($Col)
                }  
                if ($property.Gettype().IsArray) {
                    $DR.Item($property.Name) =$property.value | ConvertTo-XML -AS String -NoTypeInformation -Depth 1
                }  
               else {
                    If ($Property.Value) {
                        $DR.Item($Property.Name) = $Property.Value
                    } Else {
                        $DR.Item($Property.Name)=[DBNull]::Value
                    }
                }
            }  
            $DT.Rows.Add($DR)  
            $First = $false
        }
    } 
     
    End
    {
        Write-Output @(,($dt))
    }

}

Function Write-DataTable {
    [CmdletBinding()]
    param(
    [Parameter(Position=0, Mandatory=$true)] 
    [string]$Computername,
    [Parameter(Position=1, Mandatory=$true)] 
    [string]$Database,
    [Parameter(Position=2, Mandatory=$true)] 
    [string]$TableName,
    [Parameter(Position=3, Mandatory=$true)] 
    $Data,
    [Parameter(Position=4)] 
    [string]$Username,
    [Parameter(Position=5)] 
    [string]$Password,
    [Parameter(Position=6)] 
    [Int32]$BatchSize=50000,
    [Parameter(Position=7)] 
    [Int32]$QueryTimeout=0,
    [Parameter(Position=8)] 
    [Int32]$ConnectionTimeout=15
    )
    
    $SQLConnection = new-object System.Data.SqlClient.SQLConnection

    If ($Username) { 
        $ConnectionString = "Server={0};Database={1};User ID={2};Password={3};Trusted_Connection=False;Connect Timeout={4}" -f $Computername,$Database,$Username,$Password,$ConnectionTimeout 
    }
    Else { 
        $ConnectionString = "Server={0};Database={1};Integrated Security=True;Connect Timeout={2}" -f $Computername,$Database,$ConnectionTimeout 
    }

    $SQLConnection.ConnectionString = $ConnectionString

    Try {
        $SQLConnection.Open()
        $bulkCopy = New-Object Data.SqlClient.SqlBulkCopy -ArgumentList $SQLConnection, ([System.Data.SqlClient.SqlBulkCopyOptions]::TableLock),$Null
        $bulkCopy.DestinationTableName = $tableName
        $bulkCopy.BatchSize = $BatchSize
        $bulkCopy.BulkCopyTimeout = $QueryTimeOut
        $bulkCopy.WriteToServer($Data)        
    }
    Catch {
        Write-Error "$($TableName): $($_)"
    }
    Finally {
        $SQLConnection.Close()
    }
}

Function Get-Server {
    [CmdletBinding()]
    Param (
    [parameter(Mandatory=$False)]
    [ValidateSet('Contoso','PHX','TST','BIZ')]
    [String]
    $Domain = 'Contoso'
)     
    $Searcher = [adsisearcher]""
    If ($Domain = 'PHX')
    {
    $searchroot = [ADSI]'LDAP://DC=phx,DC=contoso,DC=com'
    }
    elseif ($Domain = 'BIZ')
    {
    $searchroot = [ADSI]'LDAP://DC=CREDITONEAPP,DC=BIZ'
    }
    elseif ($Domain = 'TST')
    {
    $searchroot = [ADSI]'LDAP://DC=CREDITONEAPP,DC=TST'
    }
    else
    {
    $searchroot = [ADSI]'LDAP://DC=contoso,DC=com'
    }

    $Searcher.SearchRoot = $searchroot
    $Searcher.Filter = "(&(objectCategory=computer)(OperatingSystem=Windows*Server*))"
    $Searcher.pagesize = 10
    $Searcher.sizelimit = 5000
    $searcher.PropertiesToLoad.Add("name") | Out-Null
    $Searcher.Sort.PropertyName='name'
    $Searcher.Sort.Direction = 'Ascending'
    $Searcher.FindAll() | ForEach-Object {$_.Properties.name}
}

Function Invoke-SQLCmd {    
    [cmdletbinding(
        DefaultParameterSetName = 'NoCred',
        SupportsShouldProcess = $True,
        ConfirmImpact = 'Low'
    )]
    Param (
        [parameter()]
        [string]$Computername,
        
        [parameter()]
        [string]$Database,    
        
        [parameter()]
        [string]$TSQL,

        [parameter()]
        [int]$ConnectionTimeout = 30,

        [parameter()]
        [int]$QueryTimeout = 120,

        [parameter()]
        [System.Collections.ICollection]$SQLParameter,

        [parameter(ParameterSetName='Cred')]
        [Alias('RunAs')]        
        [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty,

        [parameter()]
        [ValidateSet('Query','NonQuery')]
        [string]$CommandType = 'Query'
    )
    If ($PSBoundParameters.ContainsKey('Debug')) {
        $DebugPreference = 'Continue'
    }
    $PSBoundParameters.GetEnumerator() | ForEach {
        Write-Debug $_
    }
    # Region Make Connection
    Write-Verbose "Building connection string"
    $Connection = New-Object System.Data.SqlClient.SQLConnection 
    If ($PSBoundParameters.ContainsKey('Verbose')) {
        $Handler = [System.Data.SqlClient.SqlInfoMessageEventHandler] {
            Param($sender, $event) 
            Write-Verbose $event.Message -Verbose
        }
        $Connection.add_InfoMessage($Handler)
        $Connection.FireInfoMessageEventOnUserErrors=$True  
    }
    Switch ($PSCmdlet.ParameterSetName) {
        'Cred' {
            $ConnectionString = "Server={0};Database={1};User ID={2};Password={3};Trusted_Connection=False;Connect Timeout={4}" -f $Computername,
                                                                                        $Database,$Credential.Username,
                                                                                        $Credential.GetNetworkCredential().password,$ConnectionTimeout   
            Remove-Variable Credential
        }
        'NoCred' {
            $ConnectionString = "Server={0};Database={1};Integrated Security=True;Connect Timeout={2}" -f $Computername,$Database,$ConnectionTimeout                 
        }
    }   
    $Connection.ConnectionString=$ConnectionString
    Write-Verbose "Opening connection to $($Computername)"
    $Connection.Open()
    # End Region Make Connection
    # Region Initiate Query
    Write-Verbose "Initiating query -> $Tsql"
    $Command = New-Object system.Data.SqlClient.SqlCommand($Tsql,$Connection)
    If ($PSBoundParameters.ContainsKey('SQLParameter')) {
        $SqlParameter.GetEnumerator() | ForEach {
            Write-Verbose "Adding SQL Parameter: $($_.Key) with Value: $($_.Value)"
            If ($_.Value -ne $null) { 
                [void]$Command.Parameters.AddWithValue($_.Key, $_.Value) 
            }
            Else { 
                [void]$Command.Parameters.AddWithValue($_.Key, [DBNull]::Value) 
            }
        }
    }
    $Command.CommandTimeout=$QueryTimeout
    If ($PSCmdlet.ShouldProcess("Computername: $($Computername) - Database: $($Database)",'Run TSQL operation')) {
        Switch ($CommandType) {
            'Query' {
                Write-Verbose "Performing Query operation"
                $DataSet = New-Object system.Data.DataSet
                $DataAdapter = New-Object system.Data.SqlClient.SqlDataAdapter($Command)
                [void]$DataAdapter.fill($DataSet)
                $DataSet.Tables
            }
            'NonQuery' {
                Write-Verbose "Performing Non-Query operation"
                [void]$Command.ExecuteNonQuery()
            }
        }
    }
    # End Region Initiate Query    
    # Region Close connection
    Write-Verbose "Closing connection"
    $Connection.Close()        
    # End Region Close connection
}

Function Get-LocalUser {
    [Cmdletbinding()] 
    Param( 
        [Parameter()] 
        [String[]]$Computername = $Computername
    )
    Function Convert-UserFlag {
        Param ($UserFlag)
        $List = New-Object System.Collections.ArrayList
        Switch ($UserFlag) {
            ($UserFlag -BOR 0x0001)  {[void]$List.Add('SCRIPT')}
            ($UserFlag -BOR 0x0002)  {[void]$List.Add('ACCOUNTDISABLE')}
            ($UserFlag -BOR 0x0008)  {[void]$List.Add('HOMEDIR_REQUIRED')}
            ($UserFlag -BOR 0x0010)  {[void]$List.Add('LOCKOUT')}
            ($UserFlag -BOR 0x0020)  {[void]$List.Add('PASSWD_NOTREQD')}
            ($UserFlag -BOR 0x0040)  {[void]$List.Add('PASSWD_CANT_CHANGE')}
            ($UserFlag -BOR 0x0080)  {[void]$List.Add('ENCRYPTED_TEXT_PWD_ALLOWED')}
            ($UserFlag -BOR 0x0100)  {[void]$List.Add('TEMP_DUPLICATE_ACCOUNT')}
            ($UserFlag -BOR 0x0200)  {[void]$List.Add('NORMAL_ACCOUNT')}
            ($UserFlag -BOR 0x0800)  {[void]$List.Add('INTERDOMAIN_TRUST_ACCOUNT')}
            ($UserFlag -BOR 0x1000)  {[void]$List.Add('WORKSTATION_TRUST_ACCOUNT')}
            ($UserFlag -BOR 0x2000)  {[void]$List.Add('SERVER_TRUST_ACCOUNT')}
            ($UserFlag -BOR 0x10000)  {[void]$List.Add('DONT_EXPIRE_PASSWORD')}
            ($UserFlag -BOR 0x20000)  {[void]$List.Add('MNS_LOGON_ACCOUNT')}
            ($UserFlag -BOR 0x40000)  {[void]$List.Add('SMARTCARD_REQUIRED')}
            ($UserFlag -BOR 0x80000)  {[void]$List.Add('TRUSTED_FOR_DELEGATION')}
            ($UserFlag -BOR 0x100000)  {[void]$List.Add('NOT_DELEGATED')}
            ($UserFlag -BOR 0x200000)  {[void]$List.Add('USE_DES_KEY_ONLY')}
            ($UserFlag -BOR 0x400000)  {[void]$List.Add('DONT_REQ_PREAUTH')}
            ($UserFlag -BOR 0x800000)  {[void]$List.Add('PASSWORD_EXPIRED')}
            ($UserFlag -BOR 0x1000000)  {[void]$List.Add('TRUSTED_TO_AUTH_FOR_DELEGATION')}
            ($UserFlag -BOR 0x04000000)  {[void]$List.Add('PARTIAL_SECRETS_ACCOUNT')}
        }
        $List -join '; '
    }
    Function ConvertTo-SID {
        Param([byte[]]$BinarySID)
        (New-Object System.Security.Principal.SecurityIdentifier($BinarySID,0)).Value
    }
    $adsi = [ADSI]"WinNT://$Computername"
    $adsi.Children | where {$_.SchemaClassName -eq 'user'} |
    Select @{L='Computername';E={$Computername}}, @{L='Name';E={$_.Name[0]}}, 
    @{L='PasswordAge';E={("{0:N0}" -f ($_.PasswordAge[0]/86400))}}, 
    @{L='LastLogin';E={If ($_.LastLogin[0] -is [datetime]){$_.LastLogin[0]}Else{$Null}}}, 
    @{L='SID';E={(ConvertTo-SID -BinarySID $_.ObjectSID[0])}}, 
    @{L='UserFlags';E={(Convert-UserFlag -UserFlag $_.UserFlags[0])}}
}

Function Get-LocalGroup {
    [Cmdletbinding()] 
    Param( 
        [Parameter()] 
        [String[]]$Computername = $Computername
    )
    Function ConvertTo-SID {
        Param([byte[]]$BinarySID)
        (New-Object System.Security.Principal.SecurityIdentifier($BinarySID,0)).Value
    }
    Function Get-LocalGroupMember {
        Param ($Group)
        $group.Invoke('members') | ForEach {
            $_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)
        }
    }
    $adsi = [ADSI]"WinNT://$Computername"
    $adsi.Children | where {$_.SchemaClassName -eq 'group'} | 
    Select @{L='Computername';E={$Computername}},@{L='Name';E={$_.Name[0]}},
    @{L='Members';E={((Get-LocalGroupMember -Group $_)) -join '; '}},
    @{L='SID';E={(ConvertTo-SID -BinarySID $_.ObjectSID[0])}},
    @{L='GroupType';E={$GroupType[[int]$_.GroupType[0]]}}
}

Function Get-SecurityUpdate {
    [Cmdletbinding()] 
    Param( 
        [Parameter()] 
        [String[]]$Computername = $Computername
    )              
    ForEach ($Computer in $Computername){ 
        $Paths = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall","SOFTWARE\\Wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")         
        ForEach($Path in $Paths) { 
            # Create an instance of the Registry Object and open the HKLM base key 
            Try { 
                $reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Computer) 
            } Catch { 
                $_ 
                Continue 
            } 
            Try {
                # Drill down into the Uninstall key using the OpenSubKey Method 
                $regkey=$reg.OpenSubKey($Path)  
                # Retrieve an array of string that contain all the subkey names 
                $subkeys=$regkey.GetSubKeyNames()      
                # Open each Subkey and use GetValue Method to return the required values for each 
                ForEach ($key in $subkeys){   
                    $thisKey=$Path+"\\"+$key   
                    $thisSubKey=$reg.OpenSubKey($thisKey)   
                    # Prevent Objects with empty DisplayName 
                    $DisplayName = $thisSubKey.getValue("DisplayName")
                    If ($DisplayName -AND $DisplayName -match '^Update for|rollup|^Security Update|^Service Pack|^HotFix') {
                        $Date = $thisSubKey.GetValue('InstallDate')
                        If ($Date) {
                            Write-Verbose $Date 
                            $Date = $Date -replace '(\d{4})(\d{2})(\d{2})','$1-$2-$3'
                            Write-Verbose $Date 
                            $Date = Get-Date $Date
                        } 
                        If ($DisplayName -match '(?<DisplayName>.*)\((?<KB>KB.*?)\).*') {
                            $DisplayName = $Matches.DisplayName
                            $HotFixID = $Matches.KB
                        }
                        Switch -Wildcard ($DisplayName) {
                            "Service Pack*" {$Description = 'Service Pack'}
                            "Hotfix*" {$Description = 'Hotfix'}
                            "Update*" {$Description = 'Update'}
                            "Security Update*" {$Description = 'Security Update'}
                            Default {$Description = 'Unknown'}
                        }
                        # Create New Object with empty Properties 
                        $Object = [pscustomobject] @{
                            Type = $Description
                            HotFixID = $HotFixID
                            InstalledOn = $Date
                            Description = $DisplayName
                        }
                        $Object
                    } 
                }   
                $reg.Close() 
            } Catch {}                  
        }  
    }  
}

Function Get-Software {
    [OutputType('System.Software.Inventory')]
    [Cmdletbinding()] 
    Param( 
        [Parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)] 
        [String[]]$Computername=$env:COMPUTERNAME
    )         
    Begin {
    }
    Process {     
        ForEach ($Computer in $Computername){ 
            If (Test-Connection -ComputerName $Computer -Count 1 -Quiet) {
                $Paths = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall","SOFTWARE\\Wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")         
                ForEach($Path in $Paths) { 
                    Write-Verbose "Checking Path: $Path"
                    # Create an instance of the Registry Object and open the HKLM base key 
                    Try { 
                        $reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Computer,'Registry64') 
                    } Catch { 
                        Write-Error $_ 
                        Continue 
                    } 
                    # Drill down into the Uninstall key using the OpenSubKey Method 
                    Try {
                        $regkey=$reg.OpenSubKey($Path)  
                        # Retrieve an array of string that contain all the subkey names 
                        $subkeys=$regkey.GetSubKeyNames()      
                        # Open each Subkey and use GetValue Method to return the required values for each 
                        ForEach ($key in $subkeys){   
                            Write-Verbose "Key: $Key"
                            $thisKey=$Path+"\\"+$key 
                            Try {  
                                $thisSubKey=$reg.OpenSubKey($thisKey)   
                                # Prevent Objects with empty DisplayName 
                                $DisplayName = $thisSubKey.getValue("DisplayName")
                                If ($DisplayName -AND $DisplayName -notmatch '^Update for|rollup|^Security Update|^Service Pack|^HotFix') {
                                    $Date = $thisSubKey.GetValue('InstallDate')
                                    If ($Date) {
                                        Try {
                                            $Date = [datetime]::ParseExact($Date, 'yyyyMMdd', $Null)
                                        } Catch{
				                            Write-Warning "$($Computer): $_ <$($Date)>"
                                            $Date = $Null
                                        }
                                    } 
                                    # Create New Object with empty Properties 
                                    $Publisher = Try {
                                        $thisSubKey.GetValue('Publisher').Trim()
                                    } 
                                    Catch {
                                        $thisSubKey.GetValue('Publisher')
                                    }
                                    $Version = Try {
                                        # Some weirdness with trailing [char]0 on some strings
                                        $thisSubKey.GetValue('DisplayVersion').TrimEnd(([char[]](32,0)))
                                    } 
                                    Catch {
                                        $thisSubKey.GetValue('DisplayVersion')
                                    }
                                    $UninstallString = Try {
                                        $thisSubKey.GetValue('UninstallString').Trim()
                                    } 
                                    Catch {
                                        $thisSubKey.GetValue('UninstallString')
                                    }
                                    $InstallLocation = Try {
                                        $thisSubKey.GetValue('InstallLocation').Trim()
                                    } 
                                    Catch {
                                        $thisSubKey.GetValue('InstallLocation')
                                    }
                                    $InstallSource = Try {
                                        $thisSubKey.GetValue('InstallSource').Trim()
                                    } 
                                    Catch {
                                        $thisSubKey.GetValue('InstallSource')
                                    }
                                    $HelpLink = Try {
                                        $thisSubKey.GetValue('HelpLink').Trim()
                                    } 
                                    Catch {
                                        $thisSubKey.GetValue('HelpLink')
                                    }
                                    $Object = [pscustomobject]@{
                                        Computername = $Computer
                                        DisplayName = $DisplayName
                                        Version = $Version
                                        InstallDate = $Date
                                        Publisher = $Publisher
                                        UninstallString = $UninstallString
                                        InstallLocation = $InstallLocation
                                        InstallSource = $InstallSource
                                        HelpLink = $thisSubKey.GetValue('HelpLink')
                                        EstimatedSizeMB = [decimal]([math]::Round(($thisSubKey.GetValue('EstimatedSize')*1024)/1MB,2))
                                    }
                                    $Object.pstypenames.insert(0,'System.Software.Inventory')
                                    Write-Output $Object
                                }
                            } Catch {
                                Write-Warning "$Key : $_"
                            }   
                        }
                    } Catch {}   
                    $reg.Close() 
                }                  
            } Else {
                Write-Error "$($Computer): unable to reach remote system!"
            }
        } 
    } 
} 

Function Get-UserShareDACL {
    [cmdletbinding()]
    Param(
        [Parameter()]
        $Computername = $Computername                     
    )                   
    Try {    
        Write-Verbose "Computer: $($Computername)"
        # Retrieve share information from comptuer
        $Shares = Get-WmiObject Win32_LogicalShareSecuritySetting -ComputerName $Computername -ErrorAction Stop
        ForEach ($Share in $Shares) {
            $MoreShare = $Share.GetRelated('Win32_Share')
            Write-Verbose "Share: $($Share.name)"
            # Try to get the security descriptor
            $SecurityDescriptor = $Share.GetSecurityDescriptor()
            # Iterate through each descriptor
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
    # Catch any errors                
    Catch {}                                                    
}

Function Get-AdminShare {
    [cmdletbinding()]
    Param (
        $Computername = $Computername
    )
    $WMIParams = @{
        Computername = $Computername
        Class = 'Win32_Share'
        Property = 'Name', 'Path', 'Description', 'Type'
        ErrorAction = 'Stop'
        Filter = "Type='2147483651' OR Type='2147483646' OR Type='2147483647' OR Type='2147483648'"
    }
    Get-WmiObject @WMIParams | Select-Object Name, Path, Description, 
    @{L='Type';E={$ShareType[[int64]$_.Type]}}
}

Function Convert-ChassisType {
    Param ([int[]]$ChassisType)
    $List = New-Object System.Collections.ArrayList
    Switch ($ChassisType) {
        0x0001  {[void]$List.Add('Other')}
        0x0002  {[void]$List.Add('Unknown')}
        0x0003  {[void]$List.Add('Desktop')}
        0x0004  {[void]$List.Add('Low Profile Desktop')}
        0x0005  {[void]$List.Add('Pizza Box')}
        0x0006  {[void]$List.Add('Mini Tower')}
        0x0007  {[void]$List.Add('Tower')}
        0x0008  {[void]$List.Add('Portable')}
        0x0009  {[void]$List.Add('Laptop')}
        0x000A  {[void]$List.Add('Notebook')}
        0x000B  {[void]$List.Add('Hand Held')}
        0x000C  {[void]$List.Add('Docking Station')}
        0x000D  {[void]$List.Add('All in One')}
        0x000E  {[void]$List.Add('Sub Notebook')}
        0x000F  {[void]$List.Add('Space-Saving')}
        0x0010  {[void]$List.Add('Lunch Box')}
        0x0011  {[void]$List.Add('Main System Chassis')}
        0x0012  {[void]$List.Add('Expansion Chassis')}
        0x0013  {[void]$List.Add('Subchassis')}
        0x0014  {[void]$List.Add('Bus Expansion Chassis')}
        0x0015  {[void]$List.Add('Peripheral Chassis')}
        0x0016  {[void]$List.Add('Storage Chassis')}
        0x0017  {[void]$List.Add('Rack Mount Chassis')}
        0x0018  {[void]$List.Add('Sealed-Case PC')}
    }
    $List -join ', '
}

Function Get-ScheduledTask {   
    [cmdletbinding()]
    Param (    
        [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string[]]$Computername = $env:COMPUTERNAME
    )
    Begin {
        $ST = New-Object -com("Schedule.Service")
    }
    Process {
        ForEach ($Computer in $Computername) {
            Try {
                $st.Connect($Computer)
                $root=  $st.GetFolder("\")
                @($root.GetTasks(0)) | ForEach {
                    $xml = ([xml]$_.xml).task
                    [pscustomobject] @{
                        Computername = $Computer
                        Task = $_.Name
                        Author = $xml.RegistrationInfo.Author
                        RunAs = $xml.Principals.Principal.UserId                        
                        Enabled = $_.Enabled
                        State = Switch ($_.State) {
                            0 {'Unknown'}
                            1 {'Disabled'}
                            2 {'Queued'}
                            3 {'Ready'}
                            4 {'Running'}
                        }
                        LastTaskResult = Switch ($_.LastTaskResult) {
                            0x0 {"Successfully completed"}
                            0x1 {"Incorrect function called"}
                            0x2 {"File not found"}
                            0xa {"Environment is not correct"}
                            0x41300 {"Task is ready to run at its next scheduled time"}
                            0x41301 {"Task is currently running"}
                            0x41302 {"Task is disabled"}
                            0x41303 {"Task has not yet run"}
                            0x41304 {"There are no more runs scheduled for this task"}
                            0x41306 {"Task is terminated"}
                            0x00041307 {"Either the task has no triggers or the existing triggers are disabled or not set"}
                            0x00041308 {"Event triggers do not have set run times"}
                            0x80041309 {"A task's trigger is not found"}
                            0x8004130A {"One or more of the properties required to run this task have not been set"}
                            0x8004130B {"There is no running instance of the task"}
                            0x8004130C {"The Task * SCHEDuler service is not installed on this computer"}
                            0x8004130D {"The task object could not be opened"}
                            0x8004130E {"The object is either an invalid task object or is not a task object"}
                            0x8004130F {"No account information could be found in the Task * SCHEDuler security database for the task indicated"}
                            0x80041310 {"Unable to establish existence of the account specified"}
                            0x80041311 {"Corruption was detected in the Task * SCHEDuler security database"}
                            0x80041312 {"Task * SCHEDuler security services are available only on Windows NT"}
                            0x80041313 {"The task object version is either unsupported or invalid"}
                            0x80041314 {"The task has been configured with an unsupported combination of account settings and run time options"}
                            0x80041315 {"The Task * SCHEDuler Service is not running"}
                            0x80041316 {"The task XML contains an unexpected node"}
                            0x80041317 {"The task XML contains an element or attribute from an unexpected namespace"}
                            0x80041318 {"The task XML contains a value which is incorrectly formatted or out of range"}
                            0x80041319 {"The task XML is missing a required element or attribute"}
                            0x8004131A {"The task XML is malformed"}
                            0x0004131B {"The task is registered, but not all specified triggers will start the task"}
                            0x0004131C {"The task is registered, but may fail to start"}
                            0x8004131D {"The task XML contains too many nodes of the same type"}
                            0x8004131E {"The task cannot be started after the trigger end boundary"}
                            0x8004131F {"An instance of this task is already running"}
                            0x80041320 {"The task will not run because the user is not logged on"}
                            0x80041321 {"The task image is corrupt or has been tampered with"}
                            0x80041322 {"The Task * SCHEDuler service is not available"}
                            0x80041323 {"The Task * SCHEDuler service is too busy to handle your request"}
                            0x80041324 {"The Task * SCHEDuler service attempted to run the task, but the task did not run due to one of the constraints in the task definition"}
                            0x00041325 {"The Task * SCHEDuler service has asked the task to run"}
                            0x80041326 {"The task is disabled"}
                            0x80041327 {"The task has properties that are not compatible with earlier versions of Windows"}
                            0x80041328 {"The task settings do not allow the task to start on demand"}
                            Default {[string]$_}
                        }
                        Command = $xml.Actions.Exec.Command
                        Arguments = $xml.Actions.Exec.Arguments
                        StartDirectory =$xml.Actions.Exec.WorkingDirectory
                        Hidden = $xml.Settings.Hidden
                    }
                }
            } Catch {
                Write-Warning ("{0}: {1}" -f $Computer, $_.Exception.Message)
            }
        }
    }
}
# SEPVersion Check
function Get-SEPVersion {

[CmdletBinding()]
param(
[Parameter(Position=0,Mandatory=$true,HelpMessage="Name of the computer to query SEP for",
ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
[Alias('CN','__SERVER','IPAddress','Server')]
[System.String]
$ComputerName
)

begin {
# Create object to enable access to the months of the year
$DateTimeFormat = New-Object System.Globalization.DateTimeFormatInfo

# Set Registry keys to query

If((Get-WmiObject -ComputerName $ComputerName Win32_OperatingSystem).OSArchitecture -eq '32-bit')
{
$SMCKey = "SOFTWARE\\Symantec\\Symantec Endpoint Protection\\SMC"
$AVKey = "SOFTWARE\\Symantec\\Symantec Endpoint Protection\\AV"
$SylinkKey = "SOFTWARE\\Symantec\\Symantec Endpoint Protection\\SMC\\SYLINK\\SyLink"
}
Else
{
$SMCKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\SMC" 
$AVKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\AV" 
$SylinkKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\SMC\\SYLINK\\SyLink" 
}
    }


process {


try {

# Connect to Registry
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$ComputerName)

# Obtain Product Version value
$SMCRegKey = $reg.opensubkey($SMCKey)
$SEPVersion = $SMCRegKey.GetValue('ProductVersion')

# Obtain Pattern File Date Value
$AVRegKey = $reg.opensubkey($AVKey)
$AVPatternFileDate = $AVRegKey.GetValue('PatternFileDate')

# Convert PatternFileDate to readable date
$AVYearFileDate = [string]($AVPatternFileDate[0] + 1970)
$AVMonthFileDate = $DateTimeFormat.MonthNames[$AVPatternFileDate[1]]
$AVDayFileDate = [string]$AVPatternFileDate[2]
$AVFileVersionDate = $AVDayFileDate + " " + $AVMonthFileDate + " " + $AVYearFileDate

# Obtain Sylink Group value
$SylinkRegKey = $reg.opensubkey($SylinkKey)
$SylinkGroup = $SylinkRegKey.GetValue('CurrentGroup')

}

catch [System.Management.Automation.MethodInvocationException]

{
$SEPVersion = "Unable to connect to computer"
$AVFileVersionDate = ""
$SylinkGroup = ""
}

$MYObject = “” | Select-Object ComputerName,SEPProductVersion,SEPDefinitionDate,SylinkGroup
$MYObject.ComputerName = $ComputerName
$MYObject.SEPProductVersion = $SEPVersion
$MYObject.SEPDefinitionDate = $AVFileVersionDate
$MYObject.SylinkGroup = $SylinkGroup
$MYObject

}
}
# End SEPVersion Check

# Activate/License Status Check
function Get-ActivationStatus {
[CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [String[]]$DNSHostName = $env:COMPUTERNAME
    )
    process {
        try {
            $wpa = Get-WmiObject SoftwareLicensingProduct -ComputerName $DNSHostName -Filter "ApplicationID = '55c92734-d682-4d71-983e-d6ec3f16059f'" -Property LicenseStatus -ErrorAction Stop
        } catch {
            $status = New-Object ComponentModel.Win32Exception ($_.Exception.ErrorCode)
            $wpa = $null    
        }
        $out = New-Object psobject -Property @{
            ComputerName = $DNSHostName;
            Status = [string]::Empty;
        }
        if ($wpa) {
            :outer foreach($item in $wpa) {
                switch ($item.LicenseStatus) {
                    0 {$out.Status = "Unlicensed"}
                    1 {$out.Status = "Licensed"; break outer}
                    2 {$out.Status = "Out-Of-Box Grace Period"; break outer}
                    3 {$out.Status = "Out-Of-Tolerance Grace Period"; break outer}
                    4 {$out.Status = "Non-Genuine Grace Period"; break outer}
                    5 {$out.Status = "Notification"; break outer}
                    6 {$out.Status = "Extended Grace"; break outer}
                    default {$out.Status = "Unknown value"}
                }
            }
        } else {$out.Status = $status.Message}
        $out
    }
}

# End Activate/License Status Check


# End Region Helper Functions
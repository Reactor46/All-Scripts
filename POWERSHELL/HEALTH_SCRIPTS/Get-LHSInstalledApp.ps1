Function Get-LHSInstalledApp 
{
<# 
.SYNOPSIS 
    List installed applications for local or remote computers.

.DESCRIPTION
    List installed applications for local or remote computers.
    
    List both 32-bit and 64-bit applications. Note that 
    dotNet 4.0 Support for Powershell 2.0 needed.
    
    Output looks like this:
    -------------------------
    ComputerName    : N104100
    AppID           : {90120000-001A-0407-0000-0000000FF1CE}
    AppName         : Microsoft Office Outlook MUI (German) 2007
    Publisher       : Microsoft Corporation
    Version         : 12.0.6612.1000
    Architecture    : 32bit
    UninstallString : MsiExec.exe /X{90120000-001A-0407-0000-0000000FF1CE} 

.PARAMETER ComputerName
    Outputs applications for the named computer(s). 
    If you omit this parameter, the local computer is assumed.

.PARAMETER AppID
    Outputs applications with the specified application ID. 
    An application's appID is equivalent to its subkey name underneath the Uninstall registry key. 
    For Windows Installer-based applications, this is the application's product code GUID 
    (e.g. {3248F0A8-6813-11D6-A77B-00B0D0160060}). Wildcards are permitted.

.PARAMETER AppName
    Outputs applications with the specified application name. 
    The AppName is the application's name as it appears in the 
    Add/Remove Programs list. Wildcards are permitted.

.PARAMETER Publisher
    Outputs applications with the specified publisher name. 
    Wildcards are permitted

.PARAMETER Version
    Outputs applications with the specified version. 
    Wildcards are permitted.

.EXAMPLE
    PS C:\> Get-LHSInstalledApp
    
    This command outputs installed applications on the current computer.

.EXAMPLE
    PS C:\> Get-LHSInstalledApp | Select-Object AppName,Version | Sort-Object AppName
    
    This command outputs a sorted list of applications on the current computer.

.EXAMPLE
    PS C:\> Get-LHSInstalledApp wks1,wks2 -Publisher "*microsoft*"
    
    This command outputs all installed Microsoft applications on the named computers.
    *   regular expression to match any characters.

.EXAMPLE
    PS C:\> Get-LHSInstalledApp wks1,wks2 -AppName "*Office 97*" 
    
    This command outputs any Application Name that match "Office 97" on the named computers.
    *   regular expression to match any characters.

.EXAMPLE
    PS C:\> Get-Content ComputerList.txt | Get-LHSInstalledApp -AppID "{1A97CF67-FEBB-436E-BD64-431FFEF72EB8}" | Select-Object ComputerName
    
    This command outputs the computer names named in ComputerList.txt that have the specified application installed.

.EXAMPLE
    Get-LHSInstalledApp | Where-Object {-not ( $_.AppID -like "KB*") } |
    ConvertTo-CSV -Delimiter ';' -NoTypeInformation | Out-File -FilePath C:\temp\AppsInfo.csv
    Invoke-Item C:\temp\AppsInfo.csv

    Outputs all installed application except KB fixes to an CSV file and opens in Excel

.INPUTS
    System.String, you can pipe ComputerNames to this Function

.OUTPUTS
    PSObjects containing the following properties:
    
    ComputerName - computer where the application is installed
    AppID - the application's AppID
    AppName - the application's name
    Publisher - the application's publisher
    Version - the application's version
    Architecture - the application's architecture (32-bit or 64-bit)
    UninstallString - the application uninstall String

.NOTES 
    More Info:
    ==========
    Why not using Get-WmiObject
    --------------------------- 
    * Win32_Product
    At first glance, Win32_Product would appear to be one of those best solutions.
    The Win32_product class is not query optimized. 
    Queries such as “select * from Win32_Product where (name like 'Sniffer%')” 
    require WMI to use the MSI provider to enumerate all of the installed 
    products and then parse the full list sequentially to handle the “where” clause:,

        * This process initiates a consistency check of packages installed, 
            and then verifying and repairing the installations.
        * If you have an application that makes use of the Win32_Product class, 
            you should contact the vendor to get an updated version that does not use this class.

    On Windows Server 2003, Windows Vista, and newer operating systems, querying Win32_Product 
    will trigger Windows Installer to perform a consistency check to verify the health of the 
    application. This consistency check could cause a repair installation to occur. You can 
    confirm this by checking the Windows Application Event log. You will see the following 
    events each time the class is queried and for each product installed:

    Event ID: 1035
    Description: Windows Installer reconfigured the product. Product Name: <ProductName>. 
    Product Version: <VersionNumber>. Product Language: <languageID>. 
    Reconfiguration success or error status: 0.

    Event ID: 7035/7036
    Description: The Windows Installer service entered the running state.

    I would not recommend querying Win32_Product in your production environment unless you are in a maintenance window.


    * Win32Reg_AddRemovePrograms
        Win32Reg_AddRemovePrograms is not a standard Windows class. 
        This WMI class is only loaded during the installation of an SMS/SCCM client.

    What is great about Win32Reg_AddRemovePrograms is that it contains similar properties and 
    returns results noticeably quicker than Win32_Product.


    Using Registry:
    ----------------
    By default, if your process is running as a 32 bit process you will end up accessing the 32 bit "reflection" of 
    the remote system. Therefore, registry keys like HKLM\Software will actually be mapped to HKLM\Software\Wow6432Node 
    which gets very frustrating! You can access the 64 bit "reflection" via WMI, but personally I find that quite painful.

    In order to use this function, the Powershell instance must support .Net 4.0 or greater.
    Fortunately, in .NET 4, the registry class had some extra features added to it which allowed for a new 
    overload "RegistryView". Therefore, you can now specify exactly which "reflection" of the registry 
    you want to access and manipulate! No more headaches!
    


    NAME: Get-LHSInatalledApp.ps1 
    AUTHOR: Pasquale Lantella
    KEYWORDS: Registry Redirection, Installed software, Registry64, WOW6432Node,Accessing Remote x64 Registry From an x86/x32 OS Computer
    Version History:
        1.3 -- 08.04.2015 
        Added Property "InstallDate", added Function Convert-LHSStringToDate to have the InstallDate as [System.Datetime]
        Added Property "LastModified", added Function Get-RegistryKeyTimestamp from Boe Prox    
        1.4 -- 13.11.2015
        replaced  If ($Program.GetValue("InstallDate") -eq $Null)
        with  If ( [string]::IsNullOrEmpty($Program.GetValue("InstallDate")) )
        minor changes in function Convert-LHSStringToDate 
            output $date on Errors
            replaced Write-Error,Write-Warning with Write-Verbose
    
 
.LINK 
    http://poshcode.org/3186
    http://blogs.technet.com/b/heyscriptingguy/archive/2011/11/13/use-powershell-to-quickly-find-installed-software.aspx
    http://msdn.microsoft.com/en-us/library/aa393067%28VS.85%29.aspx

#Requires -Version 3.0 
#>

[cmdletbinding(DefaultParameterSetName = 'Default', ConfirmImpact = 'low')]  

Param(

    [Parameter(ParameterSetName='AppID', Position=0,Mandatory=$False,ValueFromPipeline=$True)]
    [Parameter(ParameterSetName='Default', Position=0,Mandatory=$False,ValueFromPipeline=$True)]
	[string[]] $ComputerName=$ENV:COMPUTERNAME,
    
    [Parameter(ParameterSetName='AppID', Position=1)]
    [String] $AppID = "*",
    
    [Parameter(ParameterSetName='Default', Position=1)]
    [String] $AppName = "*",
    
    [Parameter(ParameterSetName='Default', Position=2)]
    [String] $Publisher = "*",
    
    [Parameter(ParameterSetName='Default', Position=3)]
    [String] $Version = "*"

   )

BEGIN {
    ${CmdletName} = $Pscmdlet.MyInvocation.MyCommand.Name

    If (!($PsVersionTable.clrVersion.Major -ge 4)) {Write-Error "Requires .Net 4.0 support" ; Return} 



Function Get-RegistryKeyTimestamp 
{
<#
    .SYNOPSIS
        Retrieves the registry key timestamp from a local or remote system.

    .DESCRIPTION
        Retrieves the registry key timestamp from a local or remote system.

    .PARAMETER RegistryKey
        Registry key object that can be passed into function.

    .PARAMETER SubKey
        The subkey path to view timestamp.

    .PARAMETER RegistryHive
        The registry hive that you will connect to.

        Accepted Values:
        ClassesRoot
        CurrentUser
        LocalMachine
        Users
        PerformanceData
        CurrentConfig
        DynData

    .EXAMPLE
        $RegistryKey = Get-Item "HKLM:\System\CurrentControlSet\Control\Lsa"
        $RegistryKey | Get-RegistryKeyTimestamp | Format-List

        FullName      : HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Lsa
        Name          : Lsa
        LastWriteTime : 12/16/2014 10:16:35 PM

        Description
        -----------
        Displays the lastwritetime timestamp for the Lsa registry key.

    .EXAMPLE
        Get-RegistryKeyTimestamp -Computername Server1 -RegistryHive LocalMachine -SubKey 'System\CurrentControlSet\Control\Lsa' |
        Format-List

        FullName      : HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Lsa
        Name          : Lsa
        LastWriteTime : 12/17/2014 6:46:08 AM

        Description
        -----------
        Displays the lastwritetime timestamp for the Lsa registry key of the remote system.

    .INPUTS
        System.String
        Microsoft.Win32.RegistryKey

    .OUTPUTS
        Microsoft.Registry.Timestamp

    .NOTES
        Name: Get-RegistryKeyTimestamp
        Author: Boe Prox
        Version History:
            1.0 -- Boe Prox 17 Dec 2014
                -Initial Build

#Requires -Version 3.0			
#>
    [OutputType('Microsoft.Registry.Timestamp')]
    [cmdletbinding(
        DefaultParameterSetName = 'ByValue'
    )]
    Param (
        [parameter(ValueFromPipeline=$True, ParameterSetName='ByValue')]
        [Microsoft.Win32.RegistryKey]$RegistryKey,
		
        [parameter(ParameterSetName='ByPath')]
        [string]$SubKey,
		
        [parameter(ParameterSetName='ByPath')]
        [Microsoft.Win32.RegistryHive]$RegistryHive,
		
        [parameter(ParameterSetName='ByPath')]
        [string]$Computername
    )
    Begin {
        #region Create Win32 API Object
        Try {
            [void][advapi32]
        } Catch {
            #region Module Builder
            $Domain = [AppDomain]::CurrentDomain
            $DynAssembly = New-Object System.Reflection.AssemblyName('RegAssembly')
            $AssemblyBuilder = $Domain.DefineDynamicAssembly($DynAssembly, [System.Reflection.Emit.AssemblyBuilderAccess]::Run) # Only run in memory
            $ModuleBuilder = $AssemblyBuilder.DefineDynamicModule('RegistryTimeStampModule', $False)
            #endregion Module Builder
 
            #region DllImport
            $TypeBuilder = $ModuleBuilder.DefineType('advapi32', 'Public, Class')
 
            #region RegQueryInfoKey Method
            $PInvokeMethod = $TypeBuilder.DefineMethod(
                'RegQueryInfoKey', #Method Name
                [Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl', #Method Attributes
                [IntPtr], #Method Return Type
                [Type[]] @(
                    [Microsoft.Win32.SafeHandles.SafeRegistryHandle], #Registry Handle
                    [System.Text.StringBuilder], #Class Name
                    [UInt32 ].MakeByRefType(),  #Class Length
                    [UInt32], #Reserved
                    [UInt32 ].MakeByRefType(), #Subkey Count
                    [UInt32 ].MakeByRefType(), #Max Subkey Name Length
                    [UInt32 ].MakeByRefType(), #Max Class Length
                    [UInt32 ].MakeByRefType(), #Value Count
                    [UInt32 ].MakeByRefType(), #Max Value Name Length
                    [UInt32 ].MakeByRefType(), #Max Value Name Length
                    [UInt32 ].MakeByRefType(), #Security Descriptor Size           
                    [long].MakeByRefType() #LastWriteTime
                ) #Method Parameters
            )
 
            $DllImportConstructor = [Runtime.InteropServices.DllImportAttribute].GetConstructor(@([String]))
            $FieldArray = [Reflection.FieldInfo[]] @(       
                [Runtime.InteropServices.DllImportAttribute].GetField('EntryPoint'),
                [Runtime.InteropServices.DllImportAttribute].GetField('SetLastError')
            )
 
            $FieldValueArray = [Object[]] @(
                'RegQueryInfoKey', #CASE SENSITIVE!!
                $True
            )
 
            $SetLastErrorCustomAttribute = New-Object Reflection.Emit.CustomAttributeBuilder(
                $DllImportConstructor,
                @('advapi32.dll'),
                $FieldArray,
                $FieldValueArray
            )
 
            $PInvokeMethod.SetCustomAttribute($SetLastErrorCustomAttribute)
            #endregion RegQueryInfoKey Method
 
            [void]$TypeBuilder.CreateType()
            #endregion DllImport
        }
        #endregion Create Win32 API object
    }
    Process {
        #region Constant Variables
        $ClassLength = 255
        [long]$TimeStamp = $null
        #endregion Constant Variables
 
        #region Registry Key Data
        If ($PSCmdlet.ParameterSetName -eq 'ByPath') {
            #Get registry key data
            $RegistryKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegistryHive, $Computername).OpenSubKey($SubKey)
            If ($RegistryKey -isnot [Microsoft.Win32.RegistryKey]) {
                Throw "Cannot open or locate $SubKey on $Computername"
            }
        }
 
        $ClassName = New-Object System.Text.StringBuilder $RegistryKey.Name
        $RegistryHandle = $RegistryKey.Handle
        #endregion Registry Key Data
 
        #region Retrieve timestamp
        $Return = [advapi32]::RegQueryInfoKey(
            $RegistryHandle,
            $ClassName,
            [ref]$ClassLength,
            $Null,
            [ref]$Null,
            [ref]$Null,
            [ref]$Null,
            [ref]$Null,
            [ref]$Null,
            [ref]$Null,
            [ref]$Null,
            [ref]$TimeStamp
        )
        Switch ($Return) {
            0 {
               #Convert High/Low date to DateTime Object
                $LastWriteTime = [datetime]::FromFileTime($TimeStamp)
 
                #Return object
                $Object = [pscustomobject]@{
                    FullName = $RegistryKey.Name
                    Name = $RegistryKey.Name -replace '.*\\(.*)','$1'
                    LastWriteTime = $LastWriteTime
                }
                $Object.pstypenames.insert(0,'Microsoft.Registry.Timestamp')
                $Object
            }
            122 {
                Throw "ERROR_INSUFFICIENT_BUFFER (0x7a)"
            }
            Default {
                Throw "Error ($return) occurred"
            }
        }
        #endregion Retrieve timestamp
    }
} #end function Get-RegistryKeyTimestamp 

function Convert-LHSStringToDate 
{
<#
.SYNOPSIS
    Convert a System.String to System.DateTime.

.DESCRIPTION
    Convert a System.String to System.DateTime.
    convert a date and time provided as String to a System.DateTime Object.

.PARAMETER Date
    A date and time as String that will be converted to a System.DateTime Object

.PARAMETER Format
    The required format of the Parameter -Date,

.PARAMETER Culture
    like 'de-DE' for Germany or 'en-US' for United States to allow Months and Day names conversion.
    you can use Get-Culture to identify the current used culture.
    Default: independent on the current system settings

.INPUTS
    System.String, you can pipe Strings to this Function

.OUTPUTS
    System.DateTime (or $Null if the String could not be converted) to the Output Stream
    Errors and Warnings to the Error or Warning stream

.Example
    Convert-LHSStringToDate -Date '3/20/2013'

    Mittwoch, 20. März 2013 00:00:00

.Example
    Convert-LHSStringToDate -Date '20.3.2013'

    Mittwoch, 20. März 2013 00:00:00

.Example
    Convert-LHSStringToDate -Date '12:26-34' -Format 'HH-mm:ss','HH:mm-ss'
    Convert-LHSStringToDate -Date '12-26:34' -Format 'HH-mm:ss','HH:mm-ss'

    Dienstag, 9. Juli 2013 12:26:34
    Dienstag, 9. Juli 2013 12:26:34

    Because only time was provided, today’s date is used to complete the other date parts of the object.
    Also notice the Format parameter accepts array of strings, 
    allowing you to specify more than one format of the input.

.Example
    Convert-LHSStringToDate -Date '25#12#2013 22:30:00' -Format 'dd#MM#yyyy HH:mm:ss'  

    Mittwoch, 25. Dezember 2013 22:30:00

.Example
    Convert-LHSStringToDate -Date 'Thursday, July 4, 2013 12:26:34 PM' `
    -Format 'dddd, MMMM d, yyyy hh:mm:ss tt'

    Donnerstag, 4. Juli 2013 12:26:34

    Works only for English months and days names, because of [System.Globalization.CultureInfo]::InvariantCulture
    used as default for -Culture 

.Example
    Convert-LHSStringToDate -Date 'Mittwoch, 25. Dezember 2013' -Format 'dddd, dd. MMMM yyyy' -Culture 'de-DE'

    Mittwoch, 25. Dezember 2013 00:00:00

    Using -Culture we can use the language specific  months and days names

.Notes
    Info about -Format
    -------------------
    The string uses the following characters “d”,“f”,“F”,“g”,“h”,“H”,“K”,“m”,“M”,“s”,“t”,“y”,“z” 
    to define type, position and format of the input values. The type of the input value 
    (day, month, minute etc.) is defined by choosing the correct letter to represent the value 
    (day d, month M, minute m etc.), case matters here. The position is defined by placing the 
    character on the correct place in the string. The format is defined by how many times the 
    character is repeated (d to represent 1-31, dd for 01-31, dddd for Monday).

    The "/" custom format specifier represents the date separator, which is used to differentiate 
    years, months, and days. The appropriate localized date separator is retrieved from the 
    DateTimeFormatInfo.DateSeparator property of the current or specified culture.

    The ":" custom format specifier represents the time separator, which is used to differentiate 
    hours, minutes, and seconds. The appropriate localized time separator is retrieved from the 
    DateTimeFormatInfo.TimeSeparator property of the current or specified culture. 

    The "d", "f", "F", "g", "h", "H", "K", "m", "M", "s", "t", "y", "z", ":", or "/" characters
    in a format string are interpreted as custom format specifiers rather than as literal characters. 
    To prevent a character from being interpreted as a format specifier, you can precede it with 
    a backslash (\), which is the escape character. The escape character signifies that the 
    following character is a character literal that should be included in the result string unchanged.

    To include a backslash in a result string, you must escape it with another backslash (\\).

    Info about DateTime.TryParseExact() Method
    ------------------------------------------
    [System.Boolean]DateTime.TryParseExact Method (
        s as String, format as String, provider as IFormatProvider, style as DateTimeStyles, result as DateTime
    )
    Parameters:

    s
    Type: System.String

    A string containing a date and time to convert. 

    format
    Type: System.String

    The required format of s. See the Remarks section for more information. 

    provider
    Type: System.IFormatProvider

    An object that supplies culture-specific formatting information about s. 

    style
    Type: System.Globalization.DateTimeStyles

    A bitwise combination of one or more enumeration values that indicate the permitted format of s. 

    result
    Type: System.DateTime

    When this method returns, contains the DateTime value equivalent to the date and time 
    contained in s, if the conversion succeeded, or MinValue if the conversion failed. 
    The conversion fails if either the s or format parameter is null, is an empty string, 
    or does not contain a date and time that correspond to the pattern specified in format. 
    This parameter is passed uninitialized. 

.Link
    Custom Date and Time Format Strings
    http://msdn.microsoft.com/en-us/library/8kb3ddd4.aspx

    DateTime.TryParseExact Method 
    http://msdn.microsoft.com/en-us/library/ms131044.aspx

#Requires -Version 2.0
#>

[cmdletbinding()]  
[OutputType('DateTime','Null')]  

Param(

    [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$True,
        HelpMessage='Type an input String that should be converted to [Datetime].')]
    #[ValidateNotNullOrEmpty()]
	[String]$Date,

    [Parameter(Position=1)]
    [String[]]$Format = (
        'dd.MM.yyyy HH:mm:ss',
        'dd.MM.yyyy',
        'dd.M.yyyy HH:mm:ss',
        'dd.M.yyyy', 
        'MM/dd/yyyy HH:mm:ss',
        'MM/dd/yyyy', 
        'M/dd/yyyy HH:mm:ss',
        'M/dd/yyyy', 
        'yyyy-MM-dd',
		'yyyyMMdd',
        'yyyy-MM-dd-HH-mm', 
        'HH:mm:ss', 
        'HH-mm-ss'
    ),

    [Parameter(Position=2)]
    [System.Globalization.CultureInfo]$Culture = [System.Globalization.CultureInfo]::InvariantCulture

)

BEGIN {

    Set-StrictMode -Version Latest

    ${CmdletName} = $Pscmdlet.MyInvocation.MyCommand.Name

} # end BEGIN

PROCESS {
    Try
    {
        $result = New-Object DateTime
 
        $isConvertible = [DateTime]::TryParseExact(
            $Date, #the input string
            $Format, #the required format
            $Culture, 
            [System.Globalization.DateTimeStyles]::None, <#the format must be matched exactly, 
                no whitespace or anything more info here: http://msdn.microsoft.com/cs-cz/library/ms131044.aspx #>
            [ref]$result
        )
        If ($isConvertible) { Write-Output $result } 
        Else 
        { 
            Write-Verbose "The input String: ""$Date"" Could not be converted into [DateTime] using the required format: ""$Format"" !" 
            Write-Output $Date
        }
    }
    Catch
    {
        Write-Verbose $_
        Write-Output $Date
    }
} # end PROCESS

END { #Write-Verbose "Function ${CmdletName} finished." 
}

} #end function Convert-LHSStringToDate 
    
} # end BEGIN

PROCESS {
    #Write-Verbose -Message "${CmdletName}: Starting Process Block"
    ForEach ($Computer in $ComputerName) {
        Write-Verbose "`$Computer contains $Computer"
 	   	IF (Test-Connection -ComputerName $Computer -Count 2 -Quiet) {
            try {       
        
                Write-Verbose "Get Architechture Type of the system"
                $OSArch = (Get-WMIObject -ComputerName $Computer win32_operatingSystem -ErrorAction Stop).OSArchitecture
                if ($OSArch -like "*64*") {$Architectures = @("32bit","64bit")}
                else {$Architectures = @("32bit")}
                #Create an array to capture program objects.
                $arApplications = @()
                foreach ($Architecture in $Architectures){
                    #We have a 64bit machine, get the 32 bit software.
                    if ($Architecture -like "*64*"){
                        #Define the entry point to the registry.
                        $strSubKey = "SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
                        $SoftArchitecture = "32bit"
                        $RegViewEnum = [Microsoft.Win32.RegistryView]::Registry64
                    }
                    #We have a 32bit machine, use the 32bit registry provider.
                    elseif ($Architectures -notcontains "64bit"){
                        #Define the entry point to the registry.
                        $strSubKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
                        $SoftArchitecture = "32bit"
                        $RegViewEnum = [Microsoft.Win32.RegistryView]::Registry32
                    }
                    #We have "64bit" in our array, capture the 64bit software.
                    else{
                        #Define the entry point to the registry.
                        $strSubKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
                        $SoftArchitecture = "64bit"
                        $RegViewEnum = [Microsoft.Win32.RegistryView]::Registry64
                    }
        
                    Write-Verbose "Create a remote registry connection to the Computer."
                    $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer, $RegViewEnum)  
                    $RegKey = $Reg.OpenSubKey($strSubKey)
                
                    Write-Verbose "Get all subkeys that exist in the entry point."
                    $RegSubKeys = $RegKey.GetSubKeyNames()  
                
                    Write-Debug "Architecture : $Architecture"
                    Write-Debug "SoftArchitecture : $SoftArchitecture"
                    Write-Verbose "Enumerate the subkeys."
                    foreach ($SubKey in $RegSubKeys) 
                    {
                        Write-Debug "`$SubKey : $SubKey"  
                        $Program = $Reg.OpenSubKey("$strSubKey\\$SubKey")
                        
                        $strDisplayName = $Program.GetValue("DisplayName")
                        if ($strDisplayName -eq $NULL) { continue }  # skip entry if empty display name
                        
                        switch ($PsCmdlet.ParameterSetName) 
                        { 

                            "AppID" { if ((split-path $SubKey -leaf) -like $AppID) 
                                    {
                                        $RegKey = ("HKLM\$strSubKey\$SubKey").replace("\\","\")

                                        $output = new-object PSObject
                                        $output | add-member NoteProperty "ComputerName" -value $computer
                                        $output | add-member NoteProperty "RegKey" -value ($RegKey)  # useful when debugging
                                        $output | add-member NoteProperty "AppID" -value (split-path $SubKey -leaf)
                                        $output | add-member NoteProperty "AppName" -value $strDisplayName
                                        $output | add-member NoteProperty "Publisher" -value $Program.GetValue("Publisher")
                                        $output | add-member NoteProperty "Version" -value $Program.GetValue("DisplayVersion")
                                        $output | add-member NoteProperty "Architecture" -value $SoftArchitecture
                                        $output | add-member NoteProperty "UninstallString" -value $Program.GetValue("UninstallString")
                                        
                                        If ( [string]::IsNullOrEmpty($Program.GetValue("InstallDate")) )
                                        {
                                            $InstallDate = $Null
                                        }
                                        Else
                                        {
                                            $InstallDate = (Convert-LHSStringToDate -Date ($Program.GetValue("InstallDate")) -ErrorAction SilentlyContinue)
                                        }
                                        
                                        $output | add-member NoteProperty "InstallDate" -value $InstallDate
                                        $output | add-member NoteProperty "LastModified" -Value (Get-RegistryKeyTimestamp -RegistryKey $Program).LastWriteTime

                                        $output
                                    } #end if  
                            } #end "AppID"

                            "Default" { If (( $strDisplayName -like $AppName ) -and (
                                    $Program.GetValue("Publisher") -like $Publisher ) -and (
                                    $Program.GetValue("DisplayVersion") -like $Version ))
                                    {
                                        $RegKey = ("HKLM\$strSubKey\$SubKey").replace("\\","\")

                                        $output = new-object PSObject
                                        $output | add-member NoteProperty "ComputerName" -value $computer
                                        $output | add-member NoteProperty "RegKey" -value ($RegKey)  # useful when debugging
                                        $output | add-member NoteProperty "AppID" -value (split-path $SubKey -leaf)
                                        $output | add-member NoteProperty "AppName" -value $strDisplayName
                                        $output | add-member NoteProperty "Publisher" -value $Program.GetValue("Publisher")
                                        $output | add-member NoteProperty "Version" -value $Program.GetValue("DisplayVersion")
                                        $output | add-member NoteProperty "Architecture" -value $SoftArchitecture
                                        $output | add-member NoteProperty "UninstallString" -value $Program.GetValue("UninstallString")
                                        
										If ( [string]::IsNullOrEmpty($Program.GetValue("InstallDate")) )
                                        {
                                            $InstallDate = $Null
                                        }
                                        Else
                                        {
                                            $InstallDate = (Convert-LHSStringToDate -Date ($Program.GetValue("InstallDate")) -ErrorAction SilentlyContinue)
                                        }
                                        
										$output | add-member NoteProperty "InstallDate" -value $InstallDate
                                        $output | add-member NoteProperty "LastModified" -Value (Get-RegistryKeyTimestamp -RegistryKey $Program).LastWriteTime

                                        $output
                                    } #end if      
                            } #end "Default"                                     
                        } #end switch
                                  
                    } # end foreach ($SubKey in $RegSubKeys)  
                } # end foreach ($Architecture in $Architectures)        
            } Catch {
                write-error $_
            }
        } Else {
            Write-Warning "\\$Computer DO NOT reply to ping"            
        } # end IF (Test-Connection -ComputerName $Computer -count 2 -quiet)
    } # end ForEach ($Computer in $computerName)

} # end PROCESS

END { Write-Verbose "Function ${CmdletName} finished." }

} # end Function Get-LHSInatalledApp


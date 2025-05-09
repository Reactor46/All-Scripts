<#
.SYNOPSIS
    Gets the SQL Server Reporting Services objects on the target server
    and saved off the json/rdl/xml/txt files for importing into another SSRS Server
	
.DESCRIPTION
   Writes the SSRS Objects out to the "11 - SSRS" folder   
   Objects written include:
   FolderTree
   RDL Report Files
   Subscriptions
   Configuration Files
   Encryption Keys
   Folder and Report Permissions
   Subscriptions
   Shared Schedules

   
.EXAMPLE
    11_SSRS_Objects.ps1 localhost
	
.EXAMPLE
    11_SSRS_Objects.ps1 server01 sa password


.Inputs
    ServerName, [SQLUser], [SQLPassword]

.Outputs

	
.NOTES

	
.LINK
	https://github.com/gwalkey
	
#>

[CmdletBinding()]
Param(
  [string]$SQLInstance='localhost',
  [string]$myuser,
  [string]$mypass
)

# Load Common Modules and .NET Assemblies
Import-Module ".\SQLTranscriptase.psm1"

# Init
Set-StrictMode -Version latest;
[string]$BaseFolder = (Get-Item -Path ".\" -Verbose).FullName
Write-Host  -f Yellow -b Black "11 - SSRS Objects"
Write-Output("Server: [{0}]" -f $SQLInstance)

# Database Server connection check
$SQLCMD1 = "select serverproperty('productversion') as 'Version'"
try
{
    if ($mypass.Length -ge 1 -and $myuser.Length -ge 1) 
    {
        Write-Output "Trying SQL Auth"        
        $myver = ConnectSQLAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $SQLCMD1 -User $myuser -Password $mypass -ErrorAction Stop| select -ExpandProperty Version
        $serverauth="sql"
    }
    else
    {
        Write-Output "Trying Windows Auth"
		$myver = ConnectWinAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $SQLCMD1 -ErrorAction Stop | select -ExpandProperty Version
        $serverauth = "win"
    }

    if($myver -ne $null)
    {
        Write-Output ("SQL Version: {0}" -f $myver)
    }

}
catch
{
    Write-Host -f red "$SQLInstance appears offline."
    Set-Location $BaseFolder
	exit
}

[int]$ver = GetSQLNumericalVersion $myver


# ---------
# Functions
# ---------
function Format-Xml {
    param(
        ## Text of an XML document.
        [Parameter(ValueFromPipeline = $true)]
        [string[]]$Text
    )

    begin {
        $data = New-Object System.Collections.ArrayList
    }
    process {
        [void] $data.Add($Text -join "`n")
    }
    end {
        $doc=New-Object System.Xml.XmlDataDocument
        $doc.LoadXml($data -join "`n")
        $sw=New-Object System.Io.Stringwriter
        $writer=New-Object System.Xml.XmlTextWriter($sw)
        $writer.Formatting = [System.Xml.Formatting]::Indented
        $doc.WriteContentTo($writer)
        $sw.ToString()
    }
}


function Get-SSRSVersion
{
    [CmdletBinding()]
    Param(
        [string]$SQLInstance
    )

    # Get SSRS (NOT SQL Server) Version
    try
    {
        $ssrsProxy = Invoke-WebRequest -Uri "http://$SQLInstance/reportserver" -UseDefaultCredential
    }
    catch
    {
        Write-Output('Error getting SSRS version from WebService: {0}' -f $_.Exception.Message)
        throw('Error getting SSRS version from WebService: {0}' -f $_.Exception.Message)
    }
    $content = $ssrsProxy.Content
    $Regex = [Regex]::new("(?<=Microsoft SQL Server Reporting Services Version )(.*)(?=`">)") 
    $match = $regex.Match($content)
    if($Match.Success)            
    {
        $Global:SSRSFamily = 'SSRS'
        $version = $Match.Value
    }
    else
    {
        # Try PowerBI Reporting Server
        $Regex = [Regex]::new("(?<=Microsoft Power BI Report Server Version )(.*)(?=`">)") 
        $match = $regex.Match($content)
        if($Match.Success)            
        {
            $Global:SSRSFamily = 'PowerBI'
            $version = $Match.Value
        }
        else
        {
            $Global:SSRSFamily = 'SSRS'
            $version = '0.0'
        }
    }

    $NumericVersion = $version.Split('.')
    
    Write-Output($NumericVersion[0]+'.'+$NumericVersion[1])
}


# Create some CSS for help in column formatting during HTML exports
$myCSS = 
"
<style type='text/css'>
table
    {
        Margin: 0px 0px 0px 4px;
        Border: 1px solid rgb(190, 190, 190);
        Font-Family: Tahoma;
        Font-Size: 9pt;
        Background-Color: rgb(252, 252, 252);
    }
tr:hover td
    {
        Background-Color: rgb(150, 150, 220);
        Color: rgb(255, 255, 255);
    }
tr:nth-child(even)
    {
        Background-Color: rgb(242, 242, 242);
    }
th
    {
        Text-Align: Left;
        Color: rgb(150, 150, 220);
        Padding: 1px 4px 1px 4px;
    }
td
    {
        Vertical-Align: Top;
        Padding: 1px 4px 1px 4px;
    }
</style>
"

# -----------------------------------------
# 0) First, see if the SSRS Database exists
# -----------------------------------------
$sqlCMDExists=
"
IF (SELECT 1 FROM sys.databases WHERE name='ReportServer')=1 
BEGIN
	SELECT 1
end
ELSE
	SELECT 0
"
if ($serverauth -eq "sql")
{
    $DBExists = ConnectSQLAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $sqlCMDExists -User $myuser -Password $mypass -ErrorAction Stop
}
else
{
    $DBExists = ConnectWinAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $sqlCMDExists -ErrorAction Stop
}

if ($DBExists.Column1 -ne 1)
{
    Write-Output("SSRS Database not found on [{0}]" -f $SQLInstance)
    echo null > "$BaseFolder\$SQLInstance\11 - SSRS Catalog - Not found or cant connect.txt"
    Set-Location $BaseFolder
    exit
}

# ---------------------
# 1) Create output folders
# ---------------------
Write-Output ("Creating Output Paths...")
set-location $BaseFolder
$folderPath = "$BaseFolder\$sqlinstance"
$fullfolderPath = "$BaseFolder\$sqlinstance\11 - SSRS"
$fullfolderPathConfig = "$BaseFolder\$sqlinstance\11 - SSRS\Config"
$fullfolderPathRDL = "$BaseFolder\$sqlinstance\11 - SSRS\Reports"
$fullfolderPathSUB = "$BaseFolder\$sqlinstance\11 - SSRS\Subscriptions"
$fullfolderPathKey = "$BaseFolder\$sqlinstance\11 - SSRS\EncryptionKey"
$fullfolderPathFolders = "$BaseFolder\$sqlinstance\11 - SSRS\FolderTree"
$fullfolderPathUsers = "$BaseFolder\$sqlinstance\11 - SSRS\Users"
$fullfolderPathPermissions = "$BaseFolder\$sqlinstance\11 - SSRS\Permissions"
$fullfolderPathSchedules = "$BaseFolder\$sqlinstance\11 - SSRS\Schedules"


if(!(test-path -path $folderPath))
{
    mkdir $folderPath | Out-Null
}

if(!(test-path -path $fullfolderPath))
{
    mkdir $fullfolderPath | Out-Null
}

if(!(test-path -path $fullfolderPathRDL))
{
    mkdir $fullfolderPathRDL | Out-Null
}

if(!(test-path -path $fullfolderPathSUB))
{
    mkdir $fullfolderPathSUB | Out-Null
}

if(!(test-path -path $fullfolderPathKey))
{
    mkdir $fullfolderPathKey | Out-Null
}

if(!(test-path -path $fullfolderPathFolders))
{
    mkdir $fullfolderPathFolders | Out-Null
}

if(!(test-path -path $fullfolderPathUsers))
{
    mkdir $fullfolderPathUsers | Out-Null
}	

if(!(test-path -path $fullfolderPathConfig))
{
    mkdir $fullfolderPathConfig | Out-Null
}	

if(!(test-path -path $fullfolderPathPermissions))
{
    mkdir $fullfolderPathPermissions | Out-Null
}	

if(!(test-path -path $fullfolderPathSchedules))
{
    mkdir $fullfolderPathSchedules | Out-Null
}	


# ---------------------------------------------
# 2) Get SSRS Version and Supported Interfaces
# ---------------------------------------------
# Get SSRS Version from WebService HTML Home Page
try
{
    $SSRSVersion = Get-SSRSVersion $SQLInstance
}
catch
{
    Write-Output('{0}' -f $_.Exception.Message)
    exit
}

$SOAPAPIURL=''
$RESTAPIURL=''
Write-Output ("SSRS Version: {0}" -f $SSRSVersion)
Write-Output ("SSRS Version: {0}" -f $SSRSVersion) | out-file -Force -encoding ascii -FilePath "$fullfolderPath\SSRSVersion.txt"

if ($SSRSVersion -ne '0.0')
{

    # 2005
    if ($SSRSVersion -eq '9.0')
    {
        "Supports SOAP Interface 2005"| out-file -append -encoding ascii -FilePath "$fullfolderPath\SSRSVersion.txt"
        $SOAPAPIURL="http://$SQLInstance/ReportServer/ReportService2005.asmx"
        $RESTAPIURL=$null
    }

    # 2008
    if ($SSRSVersion -eq '10.0')
    {
        "Supports SOAP Interface 2005"| out-file -append -encoding ascii -FilePath "$fullfolderPath\SSRSVersion.txt"
        $SOAPAPIURL="http://$SQLInstance/ReportServer/ReportService2005.asmx"
        $RESTAPIURL=$null
    }

    # 2008 R2
    if ($SSRSVersion -eq '10.5')
    {
        "Supports SOAP Interface 2010"| out-file -append -encoding ascii -FilePath "$fullfolderPath\SSRSVersion.txt"
        $SOAPAPIURL="http://$SQLInstance/ReportServer/ReportService2010.asmx"
        $RESTAPIURL=$null
    }

    # 2012
    if ($SSRSVersion -eq '11.0')
    {
        "Supports SOAP Interface 2010"| out-file -append -encoding ascii -FilePath "$fullfolderPath\SSRSVersion.txt"
        $SOAPAPIURL="http://$SQLInstance/ReportServer/ReportService2010.asmx"
        $RESTAPIURL=$null
    }

    # 2014
    if ($SSRSVersion -eq '12.0')
    {
        "Supports SOAP Interface 2010"| out-file -append -encoding ascii -FilePath "$fullfolderPath\SSRSVersion.txt"
        $SOAPAPIURL="http://$SQLInstance/ReportServer/ReportService2010.asmx"
        $RESTAPIURL=$null
    }

    # 2016
    if ($SSRSVersion -eq '13.0')
    {
        "Supports SOAP Interface 2010"| out-file -append -encoding ascii -FilePath "$fullfolderPath\SSRSVersion.txt"
        "Supports REST Interface v1.0"| out-file -append -encoding ascii -FilePath "$fullfolderPath\SSRSVersion.txt"
        $SOAPAPIURL="http://$SQLInstance/ReportServer/ReportService2010.asmx"
        $RESTAPIURL="http://$SQLInstance/reports/api/v1.0"
    }

    # 2017
    if ($SSRSVersion -eq '14.0')
    {
        "Supports SOAP Interface 2010"| out-file -append -encoding ascii -FilePath "$fullfolderPath\SSRSVersion.txt"
        "Supports REST Interface v2.0"| out-file -append -encoding ascii -FilePath "$fullfolderPath\SSRSVersion.txt"
        $SOAPAPIURL="http://$SQLInstance/ReportServer/ReportService2010.asmx"
        $RESTAPIURL="http://$SQLInstance/reports/api/v2.0"
    }
    
    # PowerBI
    if ($SSRSVersion -eq '15.0')
    {
        "Supports SOAP Interface 2010"| out-file -append -encoding ascii -FilePath "$fullfolderPath\SSRSVersion.txt"
        "Supports REST Interface v2.0"| out-file -append -encoding ascii -FilePath "$fullfolderPath\SSRSVersion.txt"
        $SOAPAPIURL="http://$SQLInstance/ReportServer/ReportService2010.asmx"
        $RESTAPIURL="http://$SQLInstance/reports/api/v2.0"
    }
}



# --------
# 3) RDL, Shared Data Sources, Shared DataSets
# --------
# Create Output Folder Structure to mirror the SSRS ReportServer Catalog tree
Write-Output "PowerBI, RDL, Shared Data Sources, Shared DataSets..."

# SSRS Object Types
$SSRSCatalogTypes= [ordered]@{}
$SSRSCatalogTypes.Add(1,'Folder')
$SSRSCatalogTypes.Add(2,'Report')
$SSRSCatalogTypes.Add(3,'File/Resource')
$SSRSCatalogTypes.Add(4,'Linked Report')
$SSRSCatalogTypes.Add(5,'Shared Data Source')
$SSRSCatalogTypes.Add(6,'Model')
$SSRSCatalogTypes.Add(7,'')
$SSRSCatalogTypes.Add(8,'Shared DataSet')
$SSRSCatalogTypes.Add(9,'Report Part')
$SSRSCatalogTypes.Add(10,'Folder')
$SSRSCatalogTypes.Add(11,'KPI')
$SSRSCatalogTypes.Add(12,'Mobile Report')
$SSRSCatalogTypes.Add(13,'PBIX')

# Controls Permission Inheritance
$SSRSPolicyRoot= [ordered]@{}
$SSRSPolicyRoot.Add(0,'Inherited')
$SSRSPolicyRoot.Add(1,'Custom')

if ($SSRSFamily -eq 'PowerBI')
{
    $sqlCMDRDL = 
    "
    WITH ItemContentBinaries AS
    (
	    SELECT
		    ItemID,
		    ParentID,
		    Name,
            [Path],
		    [Type],
            [PolicyRoot],
		    CASE Type
		    WHEN 2 THEN 'Report'
		    WHEN 5 THEN 'Shared Data Source'
		    WHEN 7 THEN 'Report Part'
		    WHEN 8 THEN 'Shared Dataset'
		    ELSE 'Other'
		    END AS TypeDescription,
		    'BINARYXML' AS 'contentType',
		    CONVERT(varbinary(max),Content) AS Content
	    FROM ReportServer.dbo.Catalog
	    WHERE Type IN (2,5,8)

	    UNION

	    SELECT
		    c.ItemID,
		    c.ParentID,
		    c.[Name],
            c.[Path],
		    c.[Type],
            [PolicyRoot],
		    'Power BI Report' as 'TypeDescription',
		    e.ContentType,
		    CONVERT(varbinary(max),e.Content) AS Content
	    FROM 
		    ReportServer.dbo.Catalog C
	    INNER JOIN 
		    reportserver.dbo.CatalogItemExtendedContent E
	    ON
		    c.ItemID = e.ItemId
	    WHERE 
            c.Type IN (13) AND e.ContentType='CatalogItem'
        ),

    --The second CTE strips off the BOM if it exists...
    ItemContentNoBOM AS
    (
        SELECT
            ItemID,
            ParentID,
            Name,
            [Path],
            [Type],
            [PolicyRoot],
            TypeDescription,
		    ContentType,
            CASE
            WHEN LEFT(Content,3) = 0xEFBBBF
                THEN CONVERT(varbinary(max),SUBSTRING(Content,4,LEN(Content)))
            ELSE
                Content
            END AS Content
        FROM ItemContentBinaries
    )

    --The outer query gets the content in its varbinary, varchar and xml representations...
    SELECT
        ItemID,
        ParentID,
        Name,
        [Path],
        [Type],
        [PolicyRoot],
        TypeDescription,
	    ContentType,
        Content, --varbinary
        CASE
		    WHEN ContentType ='BINARYXML' THEN CONVERT(varchar(max),Content)
		    WHEN ContentType IN ('CatalogItem','DataModel','PowerBIReportDefinition') THEN null
	     END AS ContentVarchar, --varchar
        CASE 
		    WHEN ContentType ='BINARYXML' THEN CONVERT(xml,Content) 
		    WHEN ContentType IN ('CatalogItem','DataModel','PowerBIReportDefinition') THEN null
	    END AS ContentXML --xml
    FROM ItemContentNoBOM
    order by 2
    "
}
else
{
    $sqlCMDRDL = 
    "
    WITH ItemContentBinaries AS
    (
	    SELECT
		    ItemID,
		    ParentID,
		    Name,
            [Hidden],
            [Path],
		    [Type],
            [PolicyRoot],
		    CASE Type
		    WHEN 2 THEN 'Report'
		    WHEN 5 THEN 'Shared Data Source'
		    WHEN 7 THEN 'Report Part'
		    WHEN 8 THEN 'Shared Dataset'
		    ELSE 'Other'
		    END AS TypeDescription,
		    'BINARYXML' AS 'contentType',
		    CONVERT(varbinary(max),Content) AS Content
	    FROM ReportServer.dbo.Catalog
	    WHERE Type IN (2,5,7,8)
    ),

    --The second CTE strips off the BOM if it exists...
    ItemContentNoBOM AS
    (
        SELECT
            ItemID,
            ParentID,
            Name,
            [Hidden],
            [Path],
            [Type],
            [PolicyRoot],
            TypeDescription,
		    ContentType,
            CASE
            WHEN LEFT(Content,3) = 0xEFBBBF
                THEN CONVERT(varbinary(max),SUBSTRING(Content,4,LEN(Content)))
            ELSE
                Content
            END AS Content
        FROM ItemContentBinaries
    )

    --The outer query gets the content in its varbinary, varchar and xml representations...
    SELECT
        ItemID,
        ParentID,
        Name,
        [Hidden],
        [Path],
        [Type],
        [PolicyRoot],
        TypeDescription,
	    ContentType,
        Content, --varbinary
        CASE
		    WHEN ContentType ='BINARYXML' THEN CONVERT(varchar(max),Content)
		    WHEN ContentType IN ('CatalogItem','DataModel','PowerBIReportDefinition') THEN null
	     END AS ContentVarchar, --varchar
        CASE 
		    WHEN ContentType ='BINARYXML' THEN CONVERT(xml,Content) 
		    WHEN ContentType IN ('CatalogItem','DataModel','PowerBIReportDefinition') THEN null
	    END AS ContentXML --xml
    FROM ItemContentNoBOM
    order by 2
    "
}

$sqlToplevelfolders = "
--- Root
SELECT 
	[ItemId],
	[ParentID],
	'/' AS 'Path'
FROM 
	[ReportServer].[dbo].[Catalog]
where 
	Parentid is null and [Type] = 1

UNION

SELECT 
	[ItemId],
	[ParentID],
	[Path]
FROM 
	[ReportServer].[dbo].[Catalog]
where 
	Parentid is not null and [Type] = 1 AND [Name]<>'System Resources'-- AND [Path] NOT IN ('/Data Sources','/Datasets')
ORDER BY [Path]
"

# initialize arrays
$Packages = [System.Collections.ArrayList]@()
$toplevelfolders = [System.Collections.ArrayList]@()

# Connect correctly
if ($serverauth -eq "sql") 
{
    # Get Packages
    try
    {
        $Packages = ConnectSQLAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $sqlCMDRDL -User $myuser -Password $mypass -ErrorAction Stop
    }
    catch
    {
        Throw("Error Connecting to SQL: {0}" -f $error[0])
    }

    # Get Top-Level Folders
    try
    {
        $toplevelfolders = ConnectSQLAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $sqlToplevelfolders -User $myuser -Password $mypass -ErrorAction Stop
    }
    catch
    {
        Throw("Error Connecting to SQL: {0}" -f $error[0])
    }
}
else
{
     # Get Packages
    try
    {
        $Packages = ConnectWinAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $sqlCMDRDL -ErrorAction Stop
    }
    catch
    {
        Throw("Error Connecting to SQL: {0}" -f $error[0])
    }

    # Get Top-Level Folders
    try
    {
        $toplevelfolders = ConnectWinAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $sqlToplevelfolders -ErrorAction Stop
    }
    catch
    {
        Throw("Error Connecting to SQL: {0}" -f $error[0])
    }


}

# Init Exporting Json Object
$ExportableItems = [System.Collections.ArrayList]@()

# Save Each 
# Folder
# Report
# Shared Data Source
# Shared DataSet
# Power BI PBIX Report
foreach ($tlfolder in $toplevelfolders)
{
    # Create Folder Structure, Fixup forward slashes
    $myNewStruct = $fullfolderPathRDL+$tlfolder.Path
    $myNewStruct = $myNewStruct.replace('/','\')
    if(!(test-path -path $myNewStruct))
    {
        mkdir $myNewStruct | Out-Null
    }

    # Only Script out the Reports in this Folder (Report ParentID matches Folder ItemID)
    $myParentID = $tlfolder.ItemID
    Foreach ($pkg in $Packages)
    {
        if ($pkg.ParentID -eq $myParentID)
        {

            # Get the Report GUID, Name, Path (without report name)
            $pkgItemID = $pkg.ItemID.Guid
            $pkgName = $Pkg.name
            $pkgVis = $Pkg.Hidden
            $pkgPath = $Pkg.Path
            $pkgType = $pkg.Type
            $pkgPolicy = $pkg.policyRoot

            # Report RDL
            if ($pkgType -eq 2)
            {    
                $pkg.ContentXML | Out-File -Force -encoding ascii -FilePath "$myNewStruct\$pkgName.rdl"
            }

            # Shared Data Source
            if ($pkgType -eq 5)
            {    
                $pkg.ContentXML | Out-File -Force -encoding ascii -FilePath "$myNewStruct\$pkgName.rdl"
            }

            # Shared Dataset
            if ($pkgType -eq 8)
            {
                $pkg.ContentXML | Out-File -Force -encoding ascii -FilePath "$myNewStruct\$pkgName.rdl"
            }

            # Power BI Report
            if ($pkgType -eq 13)
            {
                [io.file]::WriteAllBytes("$myNewStruct\$pkgName.pbix",$pkg.Content)             
            }

            # Get Data Sources (PBIX Reports dont have Data Sources)
            if ($pkgType -in (2,5,8))
            {
                # Get all DataSources for this Item
                $sqlCMDDSrc=
                "
                SELECT 
	                [ItemID],
                    [SubscriptionID],
	                [Name],
	                [Extension],
	                [Link],
	                [CredentialRetrieval],
	                [Prompt],
	                [ConnectionString],
	                [OriginalConnectionString],
                    [OriginalConnectStringExpressionBased],
	                [UserName],
	                [Password],
	                [Flags],
	                [Version]
                FROM 
	                [ReportServer].[dbo].[DataSource]
                Where
                    ItemID ='$pkgItemID'
                "

                if ($serverauth -eq "win")
                {
                    try
                    {
                        $sqlDataSources = ConnectWinAuth -SQLInstance $SQLInstance -Database "ReportServer" -SQLExec $sqlCMDDSrc -ErrorAction Stop
                    }
                    catch
                    {
                        Write-Output("Error Connecting to SQL DataSources: {0}" -f $error[0])
                    }
                }
                else
                {
                try
                    {
                        $sqlDataSources = ConnectSQLAuth -SQLInstance $SQLInstance -Database "ReportServer" -SQLExec $sqlCMDDSrc -User $myuser -Password $mypass -ErrorAction Stop
                    }
                    catch
                    {
                        Write-Output("Error Connecting to SQL DataSources: {0}" -f $error[0])
                    }
                }
            }

            # Get all Permissions/Policies for this Item
            $sqlCMDPol=
            "
            SELECT
	            e.[Path],
                C.[UserName],
                D.[RoleName]
            FROM
                dbo.PolicyUserRole A
            INNER JOIN
                dbo.Policies B
            ON
                A.[PolicyID] = B.[PolicyID]
            INNER JOIN
                dbo.Users C
            ON 
                A.[UserID] = C.[UserID]
            INNER JOIN
                dbo.Roles D
            ON 
                A.[RoleID] = D.[RoleID]
            INNER JOIN
                dbo.Catalog E
            ON
                A.[PolicyID] = E.[PolicyID]
            WHERE
                E.ItemID='$pkgItemID'
            ORDER BY
                C.[UserName],D.[RoleName]
            "

            if ($serverauth -eq "win")
            {
                try
                {
                    $sqlPolicies = ConnectWinAuth -SQLInstance $SQLInstance -Database "ReportServer" -SQLExec $sqlCMDPol -ErrorAction Stop
                }
                catch
                {
                    Write-Output("Error Connecting to SQL Policies: {0}" -f $error[0])
                }
            }
            else
            {
            try
                {
                    $sqlPolicies = ConnectSQLAuth -SQLInstance $SQLInstance -Database "ReportServer" -SQLExec $sqlCMDPol -User $myuser -Password $mypass -ErrorAction Stop
                }
                catch
                {
                    Write-Output("Error Connecting to SQL Policies: {0}" -f $error[0])
                }
            }

            # Write Report Meta Object to JSON file
            # -------------------
            # Add Regular Scalars
            $ReportMetaData = [ordered]@{}
            $ReportMetaData.Add('ItemID',$pkgItemID)
            $ReportMetaData.Add('Name',$pkgName)
            $ReportMetaData.Add('Hidden',$pkgVis)
            $ReportMetaData.Add('Path',$pkgPath)
            $ReportMetaData.Add('Type',$pkgType)
            $ReportMetaData.Add('PolicyRoot',$pkgPolicy)

            # Init Data Source Array
            $DSObjectArray = New-Object System.Collections.Generic.List[System.Object]
            
            # Build Data Source Array
            foreach ($DS in $sqlDataSources)
            {
                $DSObject = [PSCustomObject]@{
                    ItemID = $DS.ItemID
                    SubscriptionID = $DS.SubscriptionID
                    Name = $DS.Name
                    Extension = $DS.Extension
                    Link = $DS.Link
                    CredentialRetrieval = $DS.CredentialRetrieval
                    Prompt = $DS.Prompt
                    ConnectionString = $DS.ConnectionString
                    OriginalConnectionString = $DS.OriginalConnectionString
                    OriginalConnectStringExpressionBased = $DS.OriginalConnectStringExpressionBased
                    UserName = $DS.UserName
                    Password = $DS.Password
                    Flags = $DS.Flags
                    Version  = $DS.Version
                }
                $DSObjectArray.Add($DSObject)                
            }
            # ------------------------------------
            # Add Data Source Array to Meta Object
            $ReportMetaData.Add('DataSources',$DSObjectArray)

            # Init Permissions/Policies Array
            $PolicyArray = New-Object System.Collections.Generic.List[System.Object]

            # Build Data Source Array
            foreach ($Pol in $sqlPolicies)
            {
                $PolObject = [PSCustomObject]@{
                    UserName= $Pol.UserName
                    RoleName = $Pol.RoleName                    
                }
                $PolicyArray.Add($PolObject)                
            }
            # ----------------------------------------------
            # Add Permissions/Policies Array to Meta Object
            $ReportMetaData.Add('Policies',$PolicyArray)

            # Json Object
            $ExportableItems.Add($ReportMetaData) | Out-Null

        } # Parent
    } # Items in this folder
}

$ExportableItems| ConvertTo-Json -Depth 6 | out-file "$fullfolderPathRDL\ReportPathDataSource.json" -force -Encoding ascii 

# ----------------------------
# 4) SSRS Configuration Files
# ----------------------------
Write-Output "SSRS Settings to file..."

# Get WMI Instance Name First
$WMIInstance = gwmi -class "__NAMESPACE" -namespace "root\Microsoft\SqlServer\ReportServer" -ComputerName $SQLInstance | select -ExpandProperty Name

# 2008?
[int]$wmi1 = 0
try 
{
    $junk = get-wmiobject -namespace "root\Microsoft\SQLServer\ReportServer\$WMIInstance\v10\Admin" -class MSReportServer_ConfigurationSetting -computername $SQLInstance -ErrorAction Stop | out-file -FilePath "$fullfolderPathConfig\Server_Config_Settings.txt" -encoding ascii
    $wmi1 = 10
    Write-Output "Found SSRS v10 (2008)"
}
catch
{
    #Write-Output('Error: {0}' -f $_.Exception.Message)
}

# 2012?
if ($wmi1 -eq 0)
{
    try 
    {
        get-wmiobject -namespace "root\Microsoft\SQLServer\ReportServer\$WMIInstance\v11\Admin" -class MSReportServer_ConfigurationSetting -computername $SQLInstance -ErrorAction Stop | out-file -FilePath "$fullfolderPathConfig\Server_Config_Settings.txt" -encoding ascii
        $wmi1 = 11
        Write-Output "Found SSRS v11 (2012)"        
    }
    catch
    {
        #Write-Output('Error: {0}' -f $_.Exception.Message)
    }
}

# 2014?
if ($wmi1 -eq 0)
{
    try 
    {
        get-wmiobject -namespace "root\Microsoft\SQLServer\ReportServer\$WMIInstance\v12\Admin" -class MSReportServer_ConfigurationSetting -computername $SQLInstance -ErrorAction Stop | out-file -FilePath "$fullfolderPathConfig\Server_Config_Settings.txt" -encoding ascii
        $wmi1 = 12
        Write-Output "Found SSRS v12 (2014)"        
    }
    catch
    {
        #Write-Output('Error: {0}' -f $_.Exception.Message)
    }
}

# 2016?
if ($wmi1 -eq 0)
{
    try 
    {
        get-wmiobject -namespace "root\Microsoft\SQLServer\ReportServer\$WMIInstance\v13\Admin" -class MSReportServer_ConfigurationSetting -computername $SQLInstance -ErrorAction Stop | out-file -FilePath "$fullfolderPathConfig\Server_Config_Settings.txt" -encoding ascii
        $wmi1 = 13
        Write-Output "Found SSRS v13 (2016)"        
    }
    catch
    {
        #Write-Output('Error: {0}' -f $_.Exception.Message)
    }
}

# 2017
if ($wmi1 -eq 0)
{
    try 
    {
        get-wmiobject -namespace "root\Microsoft\SqlServer\ReportServer\$WMIInstance\V14\Admin" -class MSReportServer_ConfigurationSetting -computername $SQLInstance -ErrorAction Stop | out-file -FilePath "$fullfolderPathConfig\Server_Config_Settings.txt" -encoding ascii
        $wmi1 = 14
        Write-Output "Found SSRS v14 (2017)"        
    }
    catch
    {
        #Write-Output('Error: {0}' -f $_.Exception.Message)
    }
}

# Power BI?
if ($wmi1 -eq 0)
{
    try 
    {
        get-wmiobject -namespace "root\Microsoft\SQLServer\ReportServer\$WMIInstance\v15\Admin" -class MSReportServer_ConfigurationSetting -computername $SQLInstance -ErrorAction Stop | out-file -FilePath "$fullfolderPathConfig\Server_Config_Settings.txt" -encoding ascii
        $wmi1 = 15
        Write-Output "Found SSRS v15 (Power BI)"
    }
    catch
    {
        #Write-Output('Error: {0}' -f $_.Exception.Message)
    }

}

# ------------------------------
# 5) RSReportServer.config File
# ------------------------------
# https://msdn.microsoft.com/en-us/library/ms157273.aspx

Write-Output "RSReportServer.config file..."
# 2008
$copysrc = "\\$sqlinstance\c$\Program Files\Microsoft SQL Server\MSRS10.MSSQLSERVER\Reporting Services\ReportServer\RSreportserver.config"
copy-item "\\$sqlinstance\c$\Program Files\Microsoft SQL Server\MSRS10.MSSQLSERVER\Reporting Services\ReportServer\RSreportserver.config" $fullfolderPathConfig -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

# 2008 R2
$copysrc = "\\$sqlinstance\c$\Program Files\Microsoft SQL Server\MSRS10_50.MSSQLSERVER\Reporting Services\ReportServer\RSreportserver.config"
copy-item "\\$sqlinstance\c$\Program Files\Microsoft SQL Server\MSRS10_50.MSSQLSERVER\Reporting Services\ReportServer\RSreportserver.config" $fullfolderPathConfig -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

# 2012
$copysrc = "\\$sqlinstance\c$\Program Files\Microsoft SQL Server\MSRS11.MSSQLSERVER\Reporting Services\ReportServer\RSreportserver.config"
copy-item "\\$sqlinstance\c$\Program Files\Microsoft SQL Server\MSRS11.MSSQLSERVER\Reporting Services\ReportServer\RSreportserver.config" $fullfolderPathConfig -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

# 2014
$copysrc = "\\$sqlinstance\c$\Program Files\Microsoft SQL Server\MSRS12.MSSQLSERVER\Reporting Services\ReportServer\RSreportserver.config"
copy-item "\\$sqlinstance\c$\Program Files\Microsoft SQL Server\MSRS12.MSSQLSERVER\Reporting Services\ReportServer\RSreportserver.config" $fullfolderPathConfig -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

# 2016
$copysrc = "\\$sqlinstance\c$\Program Files\Microsoft SQL Server\MSRS13.MSSQLSERVER\Reporting Services\ReportServer\RSreportserver.config"
copy-item "\\$sqlinstance\c$\Program Files\Microsoft SQL Server\MSRS13.MSSQLSERVER\Reporting Services\ReportServer\RSreportserver.config" $fullfolderPathConfig -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

# 2017
$copysrc = "\\$sqlinstance\c$\Program Files\Microsoft SQL Server Reporting Services\SSRS\ReportServer\RSreportserver.config"
copy-item "\\$sqlinstance\c$\Program Files\Microsoft SQL Server Reporting Services\SSRS\ReportServer\RSreportserver.config" $fullfolderPathConfig -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

# Power BI
$copysrc = "\\$sqlinstance\c$\Program Files\Microsoft Power BI Report Server\PBIRS\ReportServer\RSreportserver.config"
copy-item "\\$sqlinstance\c$\Program Files\Microsoft Power BI Report Server\PBIRS\ReportServer\RSreportserver.config" $fullfolderPathConfig -ErrorAction SilentlyContinue -WarningAction SilentlyContinue


# ---------------------------
# 6) Database Encryption Key
# ---------------------------
Write-Output "SSRS Encryption Key..."
$SSRSEncryptionKeyPassword = "SomeNewSecurePassword$!"

# 2008 has no WMI
if ($wmi1 -eq 10)
{
    Write-Output "SSRS 2008 - cant access Encryption key from WMI. Please use rskeymgmt.exe on server to export the key"
    New-Item "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -type file -force  |Out-Null
    Add-Content -Value "Use the rskeymgmt.exe app on the SSRS server to export the encryption key" -Path "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -Encoding Ascii
}

# We can use WMI against 2012+ SSRS Servers
$old_ErrorActionPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'

# 2012
if ($wmi1 -eq 11)
{
    try
    {
        $serverClass = get-wmiobject -namespace "root\microsoft\sqlserver\reportserver\$WMIInstance\v11\admin" -class "MSReportServer_ConfigurationSetting" -computername $SQLInstance
        if ($?)
        {
            $result = $serverClass.BackupEncryptionKey($SSRSEncryptionKeyPassword)
            $stream = [System.IO.File]::Create("$fullfolderPathKey\ssrs_master_key.snk", $result.KeyFile.Length);
            $stream.Write($result.KeyFile, 0, $result.KeyFile.Length);
            $stream.Close();
        }
        else
        {
            New-Item "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -type file -force  |Out-Null
            Add-Content -Value "Use the rskeymgmt.exe app on the SSRS server to export the encryption key" -Path "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -Encoding Ascii
            Write-Output "Error Connecting to WMI for config file (v11)"
        }
    }
    catch
    {
        New-Item "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -type file -force  |Out-Null
        Add-Content -Value "Use the rskeymgmt.exe app on the SSRS server to export the encryption key" -Path "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -Encoding Ascii        
        Write-Output "Error Connecting to WMI for config file (v11) 2"
    }
}

# 2014
if ($wmi1 -eq 12)
{
    try
    {
        $serverClass = get-wmiobject -namespace "root\microsoft\sqlserver\reportserver\$WMIInstance\v12\admin" -class "MSReportServer_ConfigurationSetting" -computername $SQLInstance
        if ($?)
        {
            $result = $serverClass.BackupEncryptionKey($SSRSEncryptionKeyPassword)
            $stream = [System.IO.File]::Create("$fullfolderPathKey\ssrs_master_key.snk", $result.KeyFile.Length);
            $stream.Write($result.KeyFile, 0, $result.KeyFile.Length);
            $stream.Close();
        }
        else
        {
            New-Item "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -type file -force  |Out-Null
            Add-Content -Value "Use the rskeymgmt.exe app on the SSRS server to export the encryption key" -Path "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -Encoding Ascii            
            Write-Output "Error Connecting to WMI for config file (v12)"
        }
    }
    catch
    {
        New-Item "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -type file -force  |Out-Null
        Add-Content -Value "Use the rskeymgmt.exe app on the SSRS server to export the encryption key" -Path "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -Encoding Ascii
        Write-Output "Error Connecting to WMI for config file (v12) 2"
    }
}

# 2016
if ($wmi1 -eq 13)
{
    try
    {
        $serverClass = get-wmiobject -namespace "root\microsoft\sqlserver\reportserver\$WMIInstance\v13\admin" -class "MSReportServer_ConfigurationSetting" -computername $SQLInstance
        if ($?)
        {
            $result = $serverClass.BackupEncryptionKey($SSRSEncryptionKeyPassword)
            $stream = [System.IO.File]::Create("$fullfolderPathKey\ssrs_master_key.snk", $result.KeyFile.Length);
            $stream.Write($result.KeyFile, 0, $result.KeyFile.Length);
            $stream.Close();
        }
        else
        {
            New-Item "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -type file -force  |Out-Null
            Add-Content -Value "Use the rskeymgmt.exe app on the SSRS server to export the encryption key" -Path "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -Encoding Ascii            
            Write-Output "Error Connecting to WMI for config file (v13)"
        }
    }
    catch
    {
        New-Item "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -type file -force  |Out-Null
        Add-Content -Value "Use the rskeymgmt.exe app on the SSRS server to export the encryption key" -Path "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -Encoding Ascii
        Write-Output "Error Connecting to WMI for config file (v13) 2"
    }
}

# 2017
if ($wmi1 -eq 14)
{
    try
    {
        $serverClass = get-wmiobject -namespace "root\Microsoft\SqlServer\ReportServer\$WMIInstance\V14\Admin" -class "MSReportServer_ConfigurationSetting" -computername $SQLInstance
        if ($?)
        {
            $result = $serverClass.BackupEncryptionKey($SSRSEncryptionKeyPassword)
            $stream = [System.IO.File]::Create("$fullfolderPathKey\ssrs_master_key.snk", $result.KeyFile.Length);
            $stream.Write($result.KeyFile, 0, $result.KeyFile.Length);
            $stream.Close();
        }
        else
        {
            New-Item "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -type file -force  |Out-Null
            Add-Content -Value "Use the rskeymgmt.exe app on the SSRS server to export the encryption key" -Path "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -Encoding Ascii            
            Write-Output "Error Connecting to WMI for config file (v14)"
        }
    }
    catch
    {
        New-Item "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -type file -force  |Out-Null
        Add-Content -Value "Use the rskeymgmt.exe app on the SSRS server to export the encryption key" -Path "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -Encoding Ascii
        Write-Output "Error Connecting to WMI for config file (v14) 2"
    }
}

# Power BI Report Server
if ($wmi1 -eq 15)
{
    try
    {
        $serverClass = get-wmiobject -namespace "root\Microsoft\SqlServer\ReportServer\$WMIInstance\V15\Admin" -class "MSReportServer_ConfigurationSetting" -computername $SQLInstance
        if ($?)
        {
            $result = $serverClass.BackupEncryptionKey($SSRSEncryptionKeyPassword)
            $stream = [System.IO.File]::Create("$fullfolderPathKey\ssrs_master_key.snk", $result.KeyFile.Length);
            $stream.Write($result.KeyFile, 0, $result.KeyFile.Length);
            $stream.Close();
        }
        else
        {
            New-Item "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -type file -force  |Out-Null
            Add-Content -Value "Use the rskeymgmt.exe app on the SSRS server to export the encryption key" -Path "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -Encoding Ascii            
            Write-Output "Error Connecting to WMI for config file (v15)"
        }
    }
    catch
    {
        New-Item "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -type file -force  |Out-Null
        Add-Content -Value "Use the rskeymgmt.exe app on the SSRS server to export the encryption key" -Path "$fullfolderPathKey\SSRS_Encryption_Key_not_exported.txt" -Encoding Ascii
        Write-Output "Error Connecting to WMI for config file (v15) 2"
    }
}

$SSRSEncryptionKeyPassword | out-file -FilePath "$fullfolderPathKey\SSRS_Encryption_Key_Password.txt"

# Reset default PS error handler - cause WMI error trapping sucks
$ErrorActionPreference = $old_ErrorActionPreference 

# ---------------------
# 7) Timed Subscriptions
# ---------------------
Write-Output "Visual Timed Subscriptions as HTML..."
$myRDLSked = 
"
select 
	distinct s.ScheduleID as 'SchID',
	c.ItemId as 'ReportID',
	c.[path] as 'Folder',
	c.[name] as 'Report',
	s.state as 'State',
	case 
		when s.RecurrenceType=6 then 'Week of Month'
		when s.RecurrenceType=5 then 'Monthly'
		when s.RecurrenceType=4 then 'Daily'
		when s.RecurrenceType=2 then 'Minute'
		when s.RecurrenceType=1 then 'AdHoc'
	end as 'RecurrenceType',
	CONVERT(VARCHAR(8), s.StartDate, 108) as 'RunTime',
	coalesce(s.WeeksInterval,'') as 'Weeks_Interval',
	coalesce(s.MinutesInterval,'') as 'Minutes_Interval',
	case when s.[month] & 1 = 1 then 'X' else '' end as 'Jan',
	case when s.[month] & 2 = 2 then 'X' else '' end as 'Feb',
	case when s.[month] & 4 = 4 then 'X' else '' end as 'Mar',
	case when s.[month] & 8 = 8 then 'X' else '' end as 'Apr',
	case when s.[month] & 16 = 16 then 'X' else '' end as 'May',
	case when s.[month] & 32 = 32 then 'X' else '' end as 'Jun',
	case when s.[month] & 64 = 64 then 'X' else '' end as 'Jul',
	case when s.[month] & 128 = 128 then 'X' else '' end as 'Aug',
	case when s.[month] & 256 = 256 then 'X' else '' end as 'Sep',
	case when s.[month] & 512 = 512 then 'X' else '' end as 'Oct',
	case when s.[month] & 1024 = 1024 then 'X' else '' end as 'Nov',
	case when s.[month] & 2048 = 2048 then 'X' else '' end as 'Dec',
	case s.MonthlyWeek
		when 1 then 'First'
		when 2 then 'Second'
		when 3 then 'Third'
		when 4 then 'Fourth'
		when 5 then 'Last'
	else ''
	End AS 'Week_of_Month',
	case when s.daysofweek & 1 = 1 then 'Sun' else '' end as 'Sun',
	case when s.daysofweek & 2 = 2 then 'Mon' else '' end as 'Mon',
	case when s.daysofweek & 4 = 4 then 'Tue' else '' end as 'Tue',
	case when s.daysofweek & 8 = 8 then 'Wed' else '' end as 'Wed',
	case when s.daysofweek & 16 = 16 then 'Thu' else '' end as 'Thu',
	case when s.daysofweek & 32 = 32 then 'Fri' else '' end as 'Fri',
	case when s.daysofweek & 64 = 64 then 'Sat' else '' end as 'Sat',
	DATEPART(hh,s.StartDate) as 'RunHour',
	case when DATEPART(hh,s.StartDate) =0 then 'X' else '' end as '00Z',
	case when DATEPART(hh,s.StartDate) =1 then 'X' else '' end as '01Z',
	case when DATEPART(hh,s.StartDate) =2 then 'X' else '' end as '02Z',
	case when DATEPART(hh,s.StartDate) =3 then 'X' else '' end as '03Z',
	case when DATEPART(hh,s.StartDate) =4 then 'X' else '' end as '04Z',
	case when DATEPART(hh,s.StartDate) =5 then 'X' else '' end as '05Z',
	case when DATEPART(hh,s.StartDate) =6 then 'X' else '' end as '06Z',
	case when DATEPART(hh,s.StartDate) =7 then 'X' else '' end as '07Z',
	case when DATEPART(hh,s.StartDate) =8 then 'X' else '' end as '08Z',
	case when DATEPART(hh,s.StartDate) =9 then 'X' else '' end as '09Z',
	case when DATEPART(hh,s.StartDate) =10 then 'X' else '' end as '10Z',
	case when DATEPART(hh,s.StartDate) =11 then 'X' else '' end as '11Z',
	case when DATEPART(hh,s.StartDate) =12 then 'X' else '' end as '12Z',
	case when DATEPART(hh,s.StartDate) =13 then 'X' else '' end as '13Z',
	case when DATEPART(hh,s.StartDate) =14 then 'X' else '' end as '14Z',
	case when DATEPART(hh,s.StartDate) =15 then 'X' else '' end as '15Z',
	case when DATEPART(hh,s.StartDate) =16 then 'X' else '' end as '16Z',
	case when DATEPART(hh,s.StartDate) =17 then 'X' else '' end as '17Z',
	case when DATEPART(hh,s.StartDate) =18 then 'X' else '' end as '18Z',
	case when DATEPART(hh,s.StartDate) =19 then 'X' else '' end as '19Z',
	case when DATEPART(hh,s.StartDate) =20 then 'X' else '' end as '20Z',
	case when DATEPART(hh,s.StartDate) =21 then 'X' else '' end as '21Z',
	case when DATEPART(hh,s.StartDate) =22 then 'X' else '' end as '22Z',
	case when DATEPART(hh,s.StartDate) =23 then 'X' else '' end as '23Z'

FROM 
	[ReportServer].[dbo].[Schedule] S	            
inner join 
	[ReportServer].[dbo].[ReportSchedule]  I
on 
	S.ScheduleID = I.ScheduleID

inner join 
	[ReportServer].[dbo].[Catalog] c
on 
	I.reportID = C.ItemID
						
order by DATEPART(hh,s.StartDate), 3, 4
"

# Run Query
if ($serverauth -eq "win")
{
    try
    {
        $Skeds = ConnectWinAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $myRDLSked -ErrorAction Stop
    }
    catch
    {
        Throw("Error Connecting to SQL: {0}" -f $error[0])
    }
}
else
{
try
    {
        $Skeds = ConnectSQLAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $myRDLSked -User $myuser -Password $mypass -ErrorAction Stop
    }
    catch
    {
        Throw("Error Connecting to SQL: {0}" -f $error[0])
    }
}


$RunTime = Get-date
$HTMLFileName = "$fullfolderPathSUB\Visual_Subscription_Schedule.html"

$Skeds | select Folder, Report, State, RecurrenceType, RunTime, Weeks_Interval, Minutes_Interval, `
Jan, Feb, Mar, Apr, May, Jun, Jul, Aug, Sep, Oct, Nov, Dec, `
Week_of_Month, Sun, Mon, Tue, Wed, Thu, Fri, Sat, RunHour,  `
00Z, 01Z, 02Z, 03Z, 04Z, 05Z, 06Z, 07Z, 08Z, 09Z, 10Z, 11Z, 12Z, 13Z, 14Z, 15Z, 16Z, 17Z, 18Z, 19Z, 20Z, 21Z, 22Z, 23Z `
| ConvertTo-Html -Head $myCSS -PostContent "<h3>Ran on : $RunTime</h3>" -CSSUri "HtmlReport.css"| Set-Content $HTMLFileName

# Script out the Create Subscription Commands
Write-Output "Timed Subscriptions as SQL..."

# Older SSRS version dont have the ReportZone Column
if ($myver -ilike '9.0*' -or $myver -ilike '10.0*' -or $myver -ilike '10.5*')
{
    $mySubs = 
    "
    USE [ReportServer];

    select 
	    'exec CreateSubscription @id='+char(39)+convert(varchar(40),S.[SubscriptionID])+char(39)+', '+
	    '@Locale=N'+char(39)+S.[Locale]+char(39)+', '+
	    '@Report_Name=N'+char(39)+R.Name+char(39)+', '+
	    '@OwnerSid='+char(39)+ '0x'+convert(varchar(max),Owner.[Sid],2)+char(39)+', '+
	    '@OwnerName=N'+char(39)+ SUSER_SNAME(Owner.[Sid])+char(39)+', '+
	    '@OwnerAuthType='+char(39)+ convert(varchar,Owner.[AuthType])+char(39)+', '+
	    '@DeliveryExtension=N'+char(39)+S.[DeliveryExtension]+char(39)+', '+
	    '@InactiveFlags='+char(39)+ convert(varchar,S.[InactiveFlags])+char(39)+', '+
	    '@ExtensionSettings=N'+char(39)+ replace(convert(varchar(max),S.[ExtensionSettings]),char(39),char(39)+char(39))+char(39)+', '+
	    '@ModifiedBySid='+char(39)+ '0x'+convert(varchar(max),Modified.[Sid],2)+char(39)+', '+
	    '@ModifiedByName=N'+char(39)+isnull(SUSER_SNAME(Modified.[Sid]),'')+char(39)+', '+
	    '@ModifiedByAuthType='+char(39)+ convert(varchar,Modified.AuthType)+char(39)+', '+
	    '@ModifiedDate='+char(39)+ convert(varchar, S.[ModifiedDate],120)+char(39)+', '+
	    '@Description=N'+char(39)+S.[Description]+char(39)+', '+
	    '@LastStatus=N'+char(39)+S.[LastStatus]+char(39)+', '+
	    '@EventType=N'+char(39)+S.[EventType]+char(39)+', '+
	    '@MatchData=N'+char(39)+ replace(convert(varchar(max),S.[MatchData]),char(34),char(39)+char(39))+char(39)+', '+
	    '@Parameters=N'+char(39)+ replace(convert(varchar(max),S.[Parameters]),char(39),char(39)+char(39))+char(39)+', '+
	    '@DataSettings=N'+char(39)+ replace(convert(varchar(max),isnull(S.[DataSettings],'')),char(39),char(39)+char(39))+char(39)+', '+
	    '@Version='+char(39)+ convert(varchar,S.[Version])+char(39) as 'ExecString'
    from
        [Subscriptions] S inner join [Catalog] CAT on S.[Report_OID] = CAT.[ItemID]
        inner join [Users] Owner on S.OwnerID = Owner.UserID
        inner join [Users] Modified on S.ModifiedByID = Modified.UserID
        left outer join [SecData] SD on CAT.PolicyID = SD.PolicyID AND SD.AuthType = Owner.AuthType
        left outer join [ActiveSubscriptions] A on S.[SubscriptionID] = A.[SubscriptionID]
	    inner join [ReportServer].[dbo].[Catalog] R on S.Report_OID = r.ItemID;
    "
}
else
{
    $mySubs = 
    "
    USE [ReportServer];

    select 
	    'exec CreateSubscription @id='+char(39)+convert(varchar(40),S.[SubscriptionID])+char(39)+', '+
	    '@Locale=N'+char(39)+S.[Locale]+char(39)+', '+
	    '@Report_Name=N'+char(39)+R.Name+char(39)+', '+
	    '@ReportZone='+char(39)+ convert(varchar,S.[ReportZone])+char(39)+', '+
	    '@OwnerSid='+char(39)+ '0x'+convert(varchar(max),Owner.[Sid],2)+char(39)+', '+
	    '@OwnerName=N'+char(39)+ SUSER_SNAME(Owner.[Sid])+char(39)+', '+
	    '@OwnerAuthType='+char(39)+ convert(varchar,Owner.[AuthType])+char(39)+', '+
	    '@DeliveryExtension=N'+char(39)+S.[DeliveryExtension]+char(39)+', '+
	    '@InactiveFlags='+char(39)+ convert(varchar,S.[InactiveFlags])+char(39)+', '+
	    '@ExtensionSettings=N'+char(39)+ replace(convert(varchar(max),S.[ExtensionSettings]),char(39),char(39)+char(39))+char(39)+', '+
	    '@ModifiedBySid='+char(39)+ '0x'+convert(varchar(max),Modified.[Sid],2)+char(39)+', '+
	    '@ModifiedByName=N'+char(39)+isnull(SUSER_SNAME(Modified.[Sid]),'')+char(39)+', '+
	    '@ModifiedByAuthType='+char(39)+ convert(varchar,Modified.AuthType)+char(39)+', '+
	    '@ModifiedDate='+char(39)+ convert(varchar, S.[ModifiedDate],120)+char(39)+', '+
	    '@Description=N'+char(39)+S.[Description]+char(39)+', '+
	    '@LastStatus=N'+char(39)+S.[LastStatus]+char(39)+', '+
	    '@EventType=N'+char(39)+S.[EventType]+char(39)+', '+
	    '@MatchData=N'+char(39)+ replace(convert(varchar(max),S.[MatchData]),char(34),char(39)+char(39))+char(39)+', '+
	    '@Parameters=N'+char(39)+ replace(convert(varchar(max),S.[Parameters]),char(39),char(39)+char(39))+char(39)+', '+
	    '@DataSettings=N'+char(39)+ replace(convert(varchar(max),isnull(S.[DataSettings],'')),char(39),char(39)+char(39))+char(39)+', '+
	    '@Version='+char(39)+ convert(varchar,S.[Version])+char(39) as 'ExecString'
    from
        [Subscriptions] S inner join [Catalog] CAT on S.[Report_OID] = CAT.[ItemID]
        inner join [Users] Owner on S.OwnerID = Owner.UserID
        inner join [Users] Modified on S.ModifiedByID = Modified.UserID
        left outer join [SecData] SD on CAT.PolicyID = SD.PolicyID AND SD.AuthType = Owner.AuthType
        left outer join [ActiveSubscriptions] A on S.[SubscriptionID] = A.[SubscriptionID]
	    inner join [ReportServer].[dbo].[Catalog] R on S.Report_OID = r.ItemID;
    "
}

# Run Query
if ($serverauth -eq "win")
{
    try
    {
        $SubCommands = ConnectWinAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $mySubs -ErrorAction Stop
    }
    catch
    {
        Throw("Error Connecting to SQL: {0}" -f $error[0])
    }
}
else
{
try
    {
        $SubCommands = ConnectSQLAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $mySubs -User $myuser -Password $mypass -ErrorAction Stop
    }
    catch
    {
        Throw("Error Connecting to SQL: {0}" -f $error[0])
    }
}

# Script Out
if ($SubCommands)
{    
    New-Item "$fullfolderPathSUB\Timed_Subscriptions.sql" -type file -force  |Out-Null

    foreach ($TSub in $SubCommands)
    {        
        $TSub.ExecString | out-file -FilePath "$fullfolderPathSUB\Timed_Subscriptions.sql" -append -encoding ascii -width 500000
    }

}



# --------------------
# 8) Folder Permissions
# --------------------
# Item-level role assignments
# System-level role assignments
# Predefined roles - https://msdn.microsoft.com/en-us/library/ms157363.aspx
# * Content Manager Role
# * Publisher Role 
# * Browser Role
# * Report Builder Role
# * My Reports Role
# * System Administrator Role
# * System User Role

Write-Output "Folder and Report Permissions to JSON..."
$sqlSecurity = "
Use ReportServer;

select 
    E.Path, 
    E.Name, 
    C.UserName, 
    D.RoleName
from 
    dbo.PolicyUserRole A
inner join 
    dbo.Policies B 
ON 
    A.PolicyID = B.PolicyID
inner join 
    dbo.Users C 
on 
    A.UserID = C.UserID
inner join 
    dbo.Roles D 
on 
    A.RoleID = D.RoleID
inner join 
    dbo.Catalog E 
on 
    A.PolicyID = E.PolicyID
Where
    E.[Name] not in ('System Resources')
order by 
    1
"

# Get Permissions
if ($serverauth -eq "win")
{
    try
    {
        $sqlPermissions = ConnectWinAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $sqlSecurity -ErrorAction Stop
    }
    catch
    {
        Write-Output("Error Connecting to SQL Getting Folder Permissions: {0}" -f $error[0])
    }
}
else
{
try
    {
        $sqlPermissions = ConnectSQLAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $sqlSecurity -User $myuser -Password $mypass -ErrorAction Stop
    }
    catch
    {
        Write-Output("Error Connecting to SQL Getting Folder Permissions: {0}" -f $error[0])
    }
}

$sqlPermissions | select Path, Name, UserName, RoleName | ConvertTo-Html -Head $myCSS -PostContent "<h3>Ran on : $RunTime</h3>" | Set-Content "$fullfolderPathPermissions\FolderandReportPermissions.html"
$sqlPermissions | select Path, Name, UserName, RoleName | ConvertTo-json -Depth 4 | out-file -FilePath "$fullfolderPathPermissions\FolderandReportPermissions.json" -Force -Encoding ascii
Export-Clixml -InputObject $sqlPermissions -Path "$fullfolderPathPermissions\FolderandReportPermissions.xml" -Force -Encoding ASCII


# ------------------------
# 9) Folder Tree Structure
# ------------------------
Write-Output "FolderTree Structure with Policies to JSON..."

# Init Exporting Json Object
$FoldersToExport = [System.Collections.ArrayList]@()

# Get all Folders only (tree structure)
$sqlFolderTree = "
Use ReportServer;

--- Get Root Folder ItemID
DECLARE @RootItemId UNIQUEIDENTIFIER

SELECT 
	@RootItemId = ItemID
FROM 
	Catalog
WHERE
	Type=1 AND ParentID IS NULL
    

--- Get all Folders (Tree Structure)
SELECT 
	ItemID,
	CASE
		WHEN [Path]='' THEN '/'
	ELSE	
		[Path]
	end as 'Path',
	PolicyRoot
FROM 
	Catalog
WHERE
	Type=1 AND [Name] NOT IN ('System Resources')
ORDER BY
	[Path]
"

# Get Permissions
if ($serverauth -eq "win")
{
    try
    {
        $sqlTree = ConnectWinAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $sqlFolderTree -ErrorAction Stop
    }
    catch
    {
        Write-Output("Error Connecting to SQL Getting FolderTree Structure: {0}" -f $error[0])
    }
}
else
{
try
    {
        $sqlTree = ConnectSQLAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $sqlFolderTree -User $myuser -Password $mypass -ErrorAction Stop
    }
    catch
    {
        Write-Output("Error Connecting to SQL Getting FolderTree Structure: {0}" -f $error[0])
    }
}


# Process each folder
foreach($Folder in $sqlTree)
{
    $FolderItemID = $folder.ItemID
    $folderPath = $folder.Path
    $folderPolicyRoot = $folder.PolicyRoot

    # Add Scalar Values to Hastable
    $FolderMetaData = [ordered]@{}
    $FolderMetaData.Add('ItemID',$folderItemID)
    $FolderMetaData.Add('Path',$folderPath)
    $FolderMetaData.Add('PolicyRoot',$folderPolicyRoot)    

    # Get Policies for each Folder
    $sqlCMDFolderPolicies = 
    "
    Use ReportServer;

    SELECT
	    e.[Path],
        C.[UserName],
        D.[RoleName]
    FROM
        dbo.PolicyUserRole A
    INNER JOIN
        dbo.Policies B
    ON
        A.[PolicyID] = B.[PolicyID]
    INNER JOIN
        dbo.Users C
    ON 
        A.[UserID] = C.[UserID]
    INNER JOIN
        dbo.Roles D
    ON 
        A.[RoleID] = D.[RoleID]
    INNER JOIN
        dbo.Catalog E
    ON
        A.[PolicyID] = E.[PolicyID]
    WHERE
        E.ItemID='$FolderItemID'
    ORDER BY 
	    c.UserName,d.RoleName
    "

    # Get Permissions
    if ($serverauth -eq "win")
    {
        try
        {
            $sqlFolderpolicies = ConnectWinAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $sqlCMDFolderPolicies -ErrorAction Stop
        }
        catch
        {
            Write-Output("Error Connecting to SQL Getting FolderTree Structure: {0}" -f $error[0])
        }
    }
    else
    {
    try
        {
            $sqlFolderpolicies = ConnectSQLAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $sqlCMDFolderPolicies -User $myuser -Password $mypass -ErrorAction Stop
        }
        catch
        {
            Write-Output("Error Connecting to SQL Getting FolderTree Structure: {0}" -f $error[0])
        }
    }

    $PolObjectArray = New-Object System.Collections.Generic.List[System.Object]

    foreach($folderpol in $sqlFolderpolicies)
    {
        $myFolderObject = [PSCustomObject]@{
            UserName=$folderpol.UserName
            RoleName=$folderpol.RoleName
        }
        $PolObjectArray.Add($myFolderObject) | out-null
    }
    $FolderMetaData.Add('Policies',$PolObjectArray)

    $FoldersToExport.Add($FolderMetaData) | Out-Null    
}

$FoldersToExport | convertto-json -Depth 4 | out-file -FilePath "$fullfolderPathFolders\FolderTreeStructure.json" -Force -Encoding ascii




# ----------------------------------
# 10) Subscriptions as JSON Document
# ----------------------------------
# Use the REST API hotness on newer versions
if ($SSRSVersion -in ('14.0','15.0'))
{
    Write-Output "Subscriptions to JSON..."
    $URI = $RESTAPIURL
    $response  = Invoke-RestMethod "$URI/Subscriptions" -Method get -UseDefaultCredentials
    $RESTsubscriptions = $response.value
    $SubsToExport = [System.Collections.ArrayList]@()

    foreach($sub in $RESTSubscriptions)
    {
        # Get Extra SubDetails
        $SubID = $Sub.ID 
        $FullSubDetails = Invoke-RestMethod "$URI/Subscriptions($SubID)" -Method get -UseDefaultCredentials
        $mySubObject = [PSCustomObject]@{
            ID     =              $sub.ID
            Owner =               $sub.Owner
            IsDataDriven=         $sub.IsDataDriven
            Description=          $sub.Description
            Report=               $sub.Report
            IsActive=             $sub.IsActive
            EventType=            $sub.EventType
            Schedule=             $sub.Schedule
            ScheduleDescription=  $FullSubDetails.ScheduleDescription
            LastRuntime =         $sub.LastRunTime
            LastStatus=           $sub.LastStatus
            ExtensionSettings=    $FullSubDetails.ExtensionSettings
            DeliveryExtension=    $sub.DeliveryExtension
            ReportParameters=     $sub.ParameterValues
        }
        $SubsToExport.Add($mySubObject) | out-null
    }
    $SubsToExport| ConvertTo-Json -Depth 6 | out-file -FilePath "$fullfolderPathSUB\SubscriptionsREST.json" -force -Encoding ascii 
}


# Use the SOAP API on all versions
Write-Output "Subscriptions to XML..."
if ($SSRSVersion -in '9.0')
{
    $ReportServerUri  = "http://$SQLInstance/ReportServer/ReportService2005.asmx"
}
else
{
    $ReportServerUri  = "http://$SQLInstance/ReportServer/ReportService2010.asmx"
}

# Get SOAP Proxy
$rs2010 = New-WebServiceProxy -Uri $ReportServerUri -UseDefaultCredential;
$type = $rs2010.GetType().Namespace
$ExtensionSettingsDataType = ($type + '.ExtensionSettings')
$ActiveStateDataType = ($type + '.ActiveState')
$ParmValueDataType = ($type + '.ParameterValue')

$SOAPSubscriptions= $rs2010.ListSubscriptions("/")
$SubsToExport2 = [System.Collections.ArrayList]@()

foreach ($sub in $SOAPSubscriptions)
{
    $extSettings = New-Object ($ExtensionSettingsDataType)
    $paramSettings = New-Object ($ParmValueDataType)
    $activeSettings = New-Object ($ActiveStateDataType)
    $desc = ""
    $status = ""
    $eventType = ""
    $matchdata = ""
    $Subproperty = $rs2010.GetSubscriptionProperties($sub.subscriptionID, [ref]$extSettings, [ref]$desc, [ref]$activeSettings, [ref]$status, [ref]$eventType, [ref]$matchData, [ref]$paramSettings)

    $mySubObject = [PSCustomObject]@{
        ID     =              $sub.SubscriptionID
        Owner =               $sub.Owner
        IsDataDriven=         $sub.IsDataDriven
        Description=          $sub.Description
        Report=               $sub.Path
        IsActive=             $true
        EventType=            $sub.EventType
        Schedule=             $matchdata
        ScheduleDescription=  $null
        LastExecuted=         $sub.LastExecuted
        Status=               $sub.Status
        ExtensionSettings=    $extSettings
        DeliveryExtension=    $extSettings.Extension
        ReportParameters=     $paramSettings
            
    }
    $SubsToExport2.Add($mySubObject) | Out-Null
}
$SubsToExport2 | export-Clixml -Path "$fullfolderPathSUB\SubscriptionsSOAP.xml" -Force -Encoding ASCII


# ----------
# 11) Users
# ----------
Write-Output "Users to JSON..."
$sqlCMDUsers=
"
SELECT 
	UserID,
	UserType,
	AuthType,
	UserName
FROM 
	[ReportServer].[dbo].[Users]
WHERE
	UserName NOT IN ('Everyone','NT AUTHORITY\SYSTEM','BUILTIN\Administrators','','sa')
Order by
    UserName
"

if ($serverauth -eq "win")
{
    try
    {
        $sqlUsers = ConnectWinAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $sqlCMDUsers -ErrorAction Stop
    }
    catch
    {
        Write-Output("Error Connecting to SQL Users: {0}" -f $error[0])
    }
}
else
{
try
    {
        $sqlUsers = ConnectSQLAuth -SQLInstance $SQLInstance -Database "master" -SQLExec $sqlCMDUsers -User $myuser -Password $mypass -ErrorAction Stop
    }
    catch
    {
        Write-Output("Error Connecting to SQL Users: {0}" -f $error[0])
    }
}

"Users" | out-file -FilePath "$fullfolderPathUsers\Users.json" -Force -Encoding ascii
$sqlUsers | select UserID, UserType, AuthType, UserName | ConvertTo-json -Depth 4 | out-file -FilePath "$fullfolderPathUsers\Users.json" -Force -Encoding ascii




# -----------------------------------------------------------------
# 12) Shared Schedules
# Used by Subscriptions as their execution time and repeat pattern
# -----------------------------------------------------------------
Write-Output "Shared Schedules to JSON..."

$sqlCMDSkeds=
"
SELECT 
	* 
FROM 
	dbo.Schedule 
WHERE 
	EventType='SharedSchedule'
"

if ($serverauth -eq "win")
{
    try
    {
        $sqlSkeds = ConnectWinAuth -SQLInstance $SQLInstance -Database "ReportServer" -SQLExec $sqlCMDSkeds -ErrorAction Stop
    }
    catch
    {
        Write-Output("Error Connecting to SQL Shared Schedules: {0}" -f $error[0])
    }
}
else
{
try
    {
        $sqlSkeds = ConnectSQLAuth -SQLInstance $SQLInstance -Database "ReportServer" -SQLExec $sqlCMDSkeds -User $myuser -Password $mypass -ErrorAction Stop
    }
    catch
    {
        Write-Output("Error Connecting to SQL Shared Schedules: {0}" -f $error[0])
    }
}

"Schedules" | out-file -FilePath "$fullfolderPathSchedules\SharedSchedules.json" -Force -Encoding ascii
$sqlSkeds | `
    select `
    	ScheduleID, `
    	Name, `
	    StartDate, `
	    Flags, `
	    NextRunTime, `
	    LastRunTime, `
	    EndDate, `
	    RecurrenceType, `
	    MinutesInterval, `
	    DaysInterval, `
	    WeeksInterval, `
	    DaysOfWeek, `
	    DaysOfMonth, `
	    Month, `
	    MonthlyWeek, `
	    State, `
	    LastRunStatus, `
	    ScheduledRunTimeout, `
	    CreatedById, `
	    EventType, `
	    EventData,` `
	    Type,` 
	    ConsistancyCheck,` 
	    Path `
    | ConvertTo-json -Depth 4 | out-file -FilePath "$fullfolderPathSchedules\SharedSchedules.json" -Force -Encoding ascii


# Return to Base
set-location $BaseFolder

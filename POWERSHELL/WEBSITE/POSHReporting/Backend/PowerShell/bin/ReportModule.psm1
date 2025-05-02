function ConvertTo-CustomHTMLTable
{
    param(
    [Parameter(
    Mandatory = $true,
    Position = 0)]
    [PSCustomObject]$InputObject,
    
    [Parameter(
    Mandatory = $true,
    Position = 1)]
    [string]$TableID,
    
    [Parameter(
    Mandatory = $false,
    Position = 2)]
    [string]$PreContent)

    begin
    {
        $template = "<table id='$TableID'><thead></thead><tbody></tbody></table>"
    }
    process
    {
        try
        {
            #Create XML object
            [xml]$HTMLTable = $template

            #Find headers
            $Headers = ($InputObject | select -First 1).psobject | 
            Select -ExpandProperty Properties | 
            Where-Object {$_.MemberType -eq "NoteProperty"} |
            Select -ExpandProperty Name

            #Adding first TR in TH
            $TR = $HTMLTable.CreateElement("tr")

            $THEAD = $HTMLTable.GetElementsByTagName("thead") | select -First 1

            $THEAD.AppendChild($TR) | Out-Null
            
            #Creating and adding TR to THEAD 
            foreach($Header in $Headers)
            {
                $TH = $HTMLTable.CreateElement("th")

                $TH.innerText = $Header

                $TR = $THEAD.GetElementsByTagName("tr") | select -First 1

                $TR.AppendChild($TH) | Out-Null
            }

            #ADDING Data rows

            #Creating and adding TR to Tbody
            $TBODY = $HTMLTable.GetElementsByTagName("tbody") | select -First 1

            foreach ($Row in $InputObject)
            {
                $TR = $HTMLTable.CreateElement("tr")
               
                foreach($Property in $Row.psobject.Properties)
                {
                    $TD = $HTMLTable.CreateElement("td")
                    
                    $Data = $Property.value

                    $TD.InnerText = $Data

                    $TR.AppendChild($TD) | Out-Null

                }#end of for each Property

                $TBODY.AppendChild($TR) | Out-Null

            }#end of for each Row in Object

            #Output the PreContent with htmltable
            $PreContent + $HTMLTable.OuterXml
        }
        catch 
        {
            Write-Error $_.Exception.Message
        }
    }

}


function Convert-XMLToHashTable
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName = $true,
                   ValueFromPipeline = $true,
                   Position=0)]
        [System.Xml.XmlElement]$XMLElement
    )

    Process
    {
        try
        {
            $HashTable = @{}

            foreach($Node in $XMLElement.ChildNodes)
            {
                if($Node."#text" -like "*;*")
                {
                    $Value = $Node."#text" -split ";"
                }
                else
                {
                    $Value = $Node."#text"
                }

                $HashTable[$Node.Name] = $Value
            }

            #Remove keys with empty values
            $KeyToRemove = @()
            
            foreach($key in $HashTable.keys)
            {
                if($HashTable[$key] -eq $null)
                {
                    $KeyToRemove += $key
                }
            }

            if( $KeyToRemove -ne $null)
            {
                foreach($key in $KeyToRemove)
                {
                    $HashTable.Remove($key)
                }
            }

            $HashTable
        }
        catch
        {
            throw
        }
    }
}

function New-SPReport 
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   Position=0)]
        [string]$HTMLTemplate,

        [Parameter(Mandatory=$true,
                   Position=1)]
        [string[]]$HTMLTables,

        [Parameter(Mandatory=$true,
                   Position=2)]
        [string]$ReportName,

        [Parameter(Mandatory=$true,
                   Position=3)]
        [string]$Path

    )

    Process
    {
        try
        {
            #Get HMTL Document as XML
            [xml]$HTML = $HTMLTemplate

            #Add Document title
            $HTML.html.head.Title = $ReportName

            #Add Report Title
            ($HTML.html.GetElementsByTagName("a") | Where-Object {$_.class -eq "topbar-title"}).Innertext = $ReportName

            #Add Tables
            $TableDiv = ($HTML.GetElementsByTagName("div") | Where-Object {$_.id -eq "tables"})
            $TableDiv.innerXml = $HTMLTables

            #Edit placeholders
            $ReportDetailsDiv = ($HTML.GetElementsByTagName("div") | Where-Object {$_.id -eq "details_Dialog"}).GetElementsByTagName("p") | Where-Object {$_.class -eq "ms-Dialog-subText"}
            $ReportDetailsDiv.InnerXml = $ReportDetailsDiv.InnerXml -replace "SERVERPLACEHOLDER", $env:COMPUTERNAME -replace "DATEPLACEHOLDER", (Get-Date).toString("dd.MM.yy HH:mm")

            #Format and save
            $XDOC = [System.Xml.Linq.XDocument]::Parse(($Html.OuterXml))
            $HTMLString = [System.Web.HttpUtility]::HtmlDecode($XDOC.toString())
            $HTMLString | Out-File $Path -Encoding UTF8

        }
        catch
        {
            throw
        }
    }
}

#Helps to group and sort object by "Status" Property
function Group-ByStatus
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $InputObject
    )

    Begin
    {
        $Ordered = @()
        $Error = @()
        $Warning = @()
        $OK = @()
    }
    Process
    {
        if($InputObject.Status -eq "Error")
        {
            $Error += $InputObject
        }
        if($InputObject.Status -eq "Warning")
        {
            $Warning += $InputObject
        }
        if($InputObject.Status -eq "OK")
        {
            $OK += $InputObject
        }
    }
    End
    {
        $Ordered += $Error
        $Ordered += $Warning
        $Ordered += $OK
        $Ordered
    }
}


#Provides compability for Port Parameter for PowerShell 2.0
function Send-MailMessage2.0
{
    [CmdletBinding()]
    Param
        (
        [Parameter(Mandatory=$false,
                   Position=0)]
        [string]$SMTPServer,

        [Parameter(Mandatory=$true,
                   Position=1)]
        [int32]$Port = 25,

        [Parameter(Mandatory=$true,
                   Position=2)]
        [string]$From,

        [Parameter(Mandatory=$true,
                   Position=3)]
        [string[]]$To,

        [Parameter(Position=4)]
        [string[]]$Cc,

        [Parameter(Position=5)]
        [string[]]$Bcc,

        [Parameter(Position=6)]
        [string]$Subject,

        [Parameter(Position=7)]
        [string]$Body,

        [Parameter(Position=8)]
        [string]$Attachment
        )

        process
        {
            try
            {
                    $SMTPClient = New-Object Net.Mail.SmtpClient($SMTPServer, $Port)

                    $SMTPClient.UseDefaultCredentials = $true

                    $Message = New-Object Net.Mail.MailMessage

                    $Message.Attachments.Add((New-Object System.Net.Mail.Attachment -ArgumentList $Attachment))

                    $Message.From = $From

                    foreach($t in $To)
                    {
                        $Message.To.Add($T)
                    }

                    if($CC)
                    {
                        foreach($c in $CC)
                        {
                            $Message.CC.Add($c)
                        }
                    }

                    if($Bcc)
                    {
                        foreach($b in $Bcc)
                        {
                            $Message.Bcc.Add($b)
                        }
                    }

                    $Message.Subject = $Subject

                    $Message.Body = $Body

                    $Message.IsBodyHtml = $true

                    $SMTPClient.Send($Message)
            }
            catch
            {
                Write-Error $_.Exception.Message
            }
        }
}
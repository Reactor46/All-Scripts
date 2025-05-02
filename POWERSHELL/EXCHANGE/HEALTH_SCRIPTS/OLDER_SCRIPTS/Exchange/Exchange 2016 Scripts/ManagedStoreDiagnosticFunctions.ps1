Set-StrictMode -Version Latest

# Functions in this file have no dependency on Exchange or store binaries,
# so may readily be used remotely.

function Get-StoreQuery
{
    <#
        .Synopsis
        Run a query against a managed store mailbox database

        .Description
        Get-StoreQuery wraps certain other cmdlets of the Exchange product, notably Get-ExchangeDiagnosticInfo, to submit a query expression to store. In general, all query expressions are of the following form:

        select (count | ([top N] column-list)) from (target) [where (expression)] [order by (order-list)]

        count         Returns only the number of rows included in the query result
        top N         Limits the maximum number of rows returned from the query
        column-list   Comma-separated list of column names, property tags, or the * character
        target        Either the special keyword Catalog, or the name of a table/table function to query
        expression    One or more logical restrictions on the rows to return from query
        order-list    Comma-separated list of column names, optionally with sort direction

        .Parameter Database
        Mailbox database to interrogate

        .Parameter Query
        Text of the query statement

        .Parameter Server
        Server to interrogate (use with ProcessId parameter)

        .Parameter ProcessId
        Worker process to interrogate (use with Server parameter)

        .Parameter Unlimited
        Overrides internal limits on the number of rows or amount of time before interrupting a query operation

        .Inputs
        None

        .Outputs
        PSObject[]

        .Example
        Get-StoreQuery -Database "db" -Query "select * from Catalog" | ft -a

        TableName             TableType      PartitionKey           Parameters     Visibility
        ---------             ---------      ------------           ----------     ----------
        AddressInfo           Table Function                                       Redacted
        Attachment            Table          MailboxPartitionNumber                Redacted
        Breadcrumbs           Table Function                                       Public
        Catalog               Table Function                                       Public
        DefaultSettings       Table Function                                       Public
        Events                Table                                                Partial
        Folder                Table          MailboxPartitionNumber                Redacted
        Globals               Table                                                Partial
        HashConversationId    Table Function                        Byte[]         Public
        HashConversationTopic Table Function                        String         Public
        Mailbox               Table                                                Partial
        Message               Table          MailboxPartitionNumber                Redacted

        Selecting from Catalog is the standard starting point for store diagnostic query. The Catalog is always visible. It lists all available targets for query, with the various columns indicating the important characteristics of each target. Target names are case-sensitive.

        TableName       Name of the target to use as part of query statements
        TableType       Indicates in general whether the content resides on disk (Table) or in memory (Table Function)
        PartitionKey    For each target, lists zero or more columns that must be included as part of a where-expression
        Parameters      For Table Functions, lists zero or more values that must accompany the query as parameters
        Visibility      Describes any restrictions on the target or the data it contains
          Public        Target and its contents have no restrictions on visibility
          Redacted      Target may be queried only if the user has elevated access
          Partial       Target may be queried, but it contains one or more columns having PII/Private restrictions
          Private       Target may never be queried

        .Example
        Get-StoreQuery -Database "db" -Query "select count from Mailbox" | ft -a

        Mailbox
        -------
              9

        Using the count keyword instead of a list of column names directs query to return only the number of rows in the result set. Selecting count is a useful and responsible practice, because it allows the user to learn about the volume of data that their query might return, without actually retrieving and serializing the content.

        .Example
        Get-StoreQuery -Database "db" -Query "select top 5 * from Mailbox" | ft -a

        MailboxNumber MailboxGuid                          OwnerADGuid                          OwnerLegacyDN
        ------------- -----------                          -----------                          -------------
                  100 f6bfe7bb-74cb-4b7a-91d9-eff17b9765da e968fb54-1554-43db-b720-7a4ee32c606c /o=First Organization/ou=Exc...
                  101 22495f47-fb5b-4d73-9251-e3a02cfa8c21 0356e767-91c4-47e0-b1a1-b9aa64d99b95 /o=First Organization/ou=Exc...
                  102 d9a163d7-a978-4209-97ba-062fa705c576 d98da2ee-e058-41a2-818a-c9545183e632 /o=First Organization/ou=Exc...
                  103 3d697f07-e8a9-4253-8473-509a4968779d 2855455c-2e77-4f7b-9626-4cdccd5bab3e /o=First Organization/ou=Use...
                  104 6add7343-dde2-47e8-9ebd-e1c87bb63185 8d2fe87f-26bc-4920-a71d-b264625aef80 /o=First Organization/ou=Use...

        Use of the "top N" expression directs query to return no more than N rows in the result set. When used, this expression must still be followed by either a comma-separated list of column names, or the use of * to indicate "all columns."

        .Example
        Get-StoreQuery -Database "db" -Query "select top 5 MailboxNumber, MailboxGuid, MessageCount, MessageSize from Mailbox" | ft -a

        MailboxNumber MailboxGuid                          MessageCount MessageSize
        ------------- -----------                          ------------ -----------
                  100 f6bfe7bb-74cb-4b7a-91d9-eff17b9765da            3       24531
                  101 22495f47-fb5b-4d73-9251-e3a02cfa8c21            1        1527
                  102 d9a163d7-a978-4209-97ba-062fa705c576            7       32589
                  103 3d697f07-e8a9-4253-8473-509a4968779d            3       23750
                  104 6add7343-dde2-47e8-9ebd-e1c87bb63185            3        5300

        In general, the * character may be used to explore a query target in order to discover what columns it contains, but the use of a comma-separated list of specific column names is encouraged, in the interests of serializing smaller volumes of data. Unlike target names, column names are usually case-insensitive.

        .Example
        Get-StoreQuery -Database "db" -Query "select top 5 MessageId from Message where MailboxPartitionNumber = 100" | ft -a *

        MessageId
        ---------
        0x82167718D6CDB04F8D115F9D1877CEC600000000001E00000100
        0x82167718D6CDB04F8D115F9D1877CEC600000001A30600000100
        0x82167718D6CDB04F8D115F9D1877CEC600000001A30700000100
        0x82167718D6CDB04F8D115F9D1877CEC600000001A6EE00000100
        0x82167718D6CDB04F8D115F9D1877CEC6000001DD2D4A00000100

        The PartitionKey value for a query target (Table or Table Function) may list the name(s) of one or more columns. For example, the PartitionKey value for the Message table is "MailboxPartitionNumber." This value indicates that any query against the Message table is required to include a where-clause which provides a specific value for MailboxPartitionNumber. To omit that expression is an error. For this reason, many investigations begin at the Mailbox table, to learn the necessary value(s) of MailboxNumber.

        .Example
        Get-StoreQuery -Database "db" -Query "select top 5 MessageId from Message where MailboxPartitionNumber = 100 and FolderId = 0x82167718D6CDB04F8D115F9D1877CEC600000000000E00000100" | ft -a *

        MessageId
        ---------
        0x82167718D6CDB04F8D115F9D1877CEC600000001A6EE00000100

        In general, string parameters that are part of a where-clause should be enclosed in quotes (single or double). Numeric arguments will not be enclosed in quotes. Some parameters, such as message IDs and folder IDs, are byte arrays that are expressed as hex strings with a 0x prefix. These hex strings may usually be copied from the result of one query and pasted into the where-clause of another, for purposes such as finding all messages within a specific folder.

        .Example
        Get-StoreQuery -Database "db" -Query "select MessageId from Message where MailboxPartitionNumber = 102 and (MessageDocumentId > 115 or FolderId = 0xD00907F04B500F41BB51AAB663B0C5A300000000000100000100)" | ft -a *

        MessageId
        ---------
        0xD00907F04B500F41BB51AAB663B0C5A300000001A2F900000100
        0xD00907F04B500F41BB51AAB663B0C5A300000001A2FA00000100
        0xD00907F04B500F41BB51AAB663B0C5A3000001DD482D00000100
        0xD00907F04B500F41BB51AAB663B0C5A300001260CC8B00000100
        0xD00907F04B500F41BB51AAB663B0C5A300001260D07400000100
        0xD00907F04B500F41BB51AAB663B0C5A300001260D07500000100

        The store query syntax for where-clauses supports parenthetical expressions and conventional comparison operators.

        =   Equal
        !=  Not Equal
        >   Greater Than
        >=  Greater Than or Equal
        <   Less Than
        <=  Less Than or Equal

        .Example
        Get-StoreQuery -Database "db" -Query "select DisplayName, FolderId from Folder where MailboxPartitionNumber = 100 and DisplayName like '%box%'" | ft -a *

        DisplayName           FolderId
        -----------           --------
        Outbox                0x82167718D6CDB04F8D115F9D1877CEC600000000000D00000100
        Inbox                 0x82167718D6CDB04F8D115F9D1877CEC600000000000E00000100
        MailboxAuditLogSearch 0x82167718D6CDB04F8D115F9D1877CEC600000000001800000100

        For string columns, query also supports "like" and "not like" keywords for matching partial strings. For wildcard matches of multiple characters, the syntax uses the SQL-like % character instead of the asterisk (*) character.

        .Example
        Get-StoreQuery -Database "db" -Query "select MessageClass from Message where MailboxPartitionNumber = 103 and MessageClass != null" | ft -a *

        MessageClass
        ------------
        IPM.Note
        IPM.ExtendedRule.Message
        IPM.Configuration.OWA.UserOptions
        IPM.Configuration.WorkHours
        IPM.Microsoft.WunderBar.Link
        IPM.Microsoft.WunderBar.Link
        IPM.Microsoft.WunderBar.Link
        IPM.Microsoft.WunderBar.Link
        IPM.Microsoft.WunderBar.Link
        IPM.Microsoft.WunderBar.Link
        IPM.Configuration.TargetFolderMRU
        IPM.Note
        IPM.Microsoft.MRM.Log

        Because columns may be null (have no value), query also allows the use of the keyword "null" as a parameter value.

        .Example
        Get-StoreQuery -Database "db" -Query "select p0037001F from Message where MailboxPartitionNumber = 102 and p0037001F != null" | ft -a *

        p0037001F
        ---------
        gal.grxml.gz_en-US_1.0
        61f7a19a-5172-4b2e-b5ba-52c4a12f5f30.grxml.gz_en-US_1.0
        User.xml_1
        DistributionList.xml_1

        Only a fraction of all property values are directly represented by column names on tables or table functions. Many more are retrieved by property tags (ptags), so query supports a ptag syntax. A ptag takes the form PNNNNTTTT, where the character P is a constant prefix, NNNN is a property number in hex, and TTTT is a property type in hex. For example, p0037001f is the ptag for Unicode Subject. Theoretically any property may be retrieved (within constraints imposed by data restrictions), including named properties, although the process for getting a named property ptag includes additional steps:

        1. Identify the mailbox number from which the named property will be selected.
        2. Examine the ExtendedPropertyNameMapping for that mailbox.
        3. Identify the property number by Guid and ID/Name as appropriate.
        4. Compose the ptag from the property number and property type.

        ptags work with select-lists, where-clauses, and order-by lists.

        .Example
        Get-StoreQuery -Database "db" -Query "select top 10 DisplayName from Folder where MailboxPartitionNumber = 100 order by DisplayName desc" | ft -a *

        DisplayName
        -----------
        Views
        Versions
        Top of Information Store
        Tasks
        Spooler Queue
        Shortcuts
        Sent Items
        Schedule
        Recoverable Items
        Purges

        Query provides an order-by syntax for sorting results, although sorting is often performed by scripts as a post-processing activity. An order-by list may include one or more column/property names, separated by commas. Each column/property name may include the suffix "asc" or "desc" to indicate ascending/descending sort direction (in the absence of any suffix, ascending is the default). The order-by list need not include the same columns/properties as the select-list.

        .Example
        Get-StoreQuery -Database "db" -Query "select MessageDocumentId, GetTopPropertySizes(PropertyBlob, OffPagePropertyBlob) from Message where MailboxPartitionNumber = 102" | ft -a *

        MessageDocumentId PropertyName1 PropertySize1 PropertyName2 PropertySize2 PropertyName3 PropertySize3
        ----------------- ------------- ------------- ------------- ------------- ------------- -------------
                        1 3FF9:Binary             137 4023:Unicode            108 8147:Unicode             92
                        2 NULL                   NULL NULL                   NULL NULL                   NULL
                        3 0E99:Binary             256 3FF9:Binary             137 4023:Unicode            108
                      103 7C07:Binary             291 3FF9:Binary             137 4023:Unicode            108
                      104 3FF9:Binary             137 4023:Unicode            108 1035:Unicode             78
                     1104 684E:Binary             179 3FF9:Binary             137 4023:Unicode            108
                     1105 684E:Binary             179 3FF9:Binary             137 4023:Unicode            108
                     1106 684E:Binary             179 3FF9:Binary             137 4023:Unicode            108
                     1107 3FF9:Binary             137 4023:Unicode            108 1035:Unicode             78

        The GetTopPropertySizes query processor accepts one or more property blob columns as arguments. It produces additional columns to indicate the identity and size (in bytes) of the largest properties found within the blob value(s) for each row.

        Use the Processors table function to see the list of all supported query processors and a brief description of each.

        .Example
        Get-StoreQuery -Database "db" -Query "select MailboxNumber, GetColumnSizes(DisplayName, PropertyBlob) from Mailbox" | ft -a *

        MailboxNumber DisplayName_Size PropertyBlob_Size
        ------------- ---------------- -----------------
                  100               36               423
                  101              102               165
                  102               90               483
                  103               90               503
                  104               90               461
                  105              128               192
                  106              128               192
                  107               36               428

        The GetColumnSizes query processor accepts one or more columns as arguments. For each argument, it produces the size (in bytes) of the column value for each row.

        Use the Processors table function to see the list of all supported query processors and a brief description of each.

        .Example
        Get-StoreQuery -Database "db" -Query "select top 1 *, -PropertyBlob, -OffPagePropertyBlob from Message where MailboxPartitionNumber = 104" | fl

        MailboxPartitionNumber     : 104
        MessageDocumentId          : 1
        MessageId                  : 0x68E6C1FC8DDA734098CFC121B33790F300000000001A00000100
        FolderId                   : 0x68E6C1FC8DDA734098CFC121B33790F300000000000E00000100
        LcnCurrent                 : 0x68E6C1FC8DDA734098CFC121B33790F300000000007A00000100
        VersionHistory             : 0x1668E6C1FC8DDA734098CFC121B33790F300000000007A
        GroupCns                   : NULL
        LastModificationTime       : 3/12/2013 8:38:01 PM
        LcnReadUnread              : 0x68E6C1FC8DDA734098CFC121B33790F300000000007A00000100
        SourceKey                  : NULL
        ChangeKey                  : NULL
        Size                       : 1276
        RecipientList              : NULL
        LargePropertyValueBlob     : NULL
        SubobjectsBlob             : NULL
        IsHidden                   : False
        IsRead                     : True
        ...

        Store query also supports Subtraction Columns, which are column names prefixed with a minus (-) sign. This expression has the effect of removing a column from the result set. The Subtraction Column prefix works with columns in the select list, and with with columns used as arguments to a processor. When combined with *, this syntax can specifically discard columns whose values are not of interest, are very large, are not suitable as processor arguments, etc.

        .Example
        Get-StoreQuery -Database "db" -Query "select MessageDocumentId, GetTopColumnSizes(*, -PropertyBlob, -OffPagePropertyBlob) from Message where MailboxPartitionNumber = 104" | ft -a *

        MessageDocumentId ColumnName1                ColumnSize1 ColumnName2   ColumnSize2 ColumnName3 ColumnSize3
        ----------------- -----------                ----------- -----------   ----------- ----------- -----------
                        1 MessageId                           26 FolderId               26 LcnCurrent           26
                        2 ConversationMembers                 60 MessageId              26 FolderId             26
                        3 MessageClass                        48 MessageId              26 FolderId             26
                      103 MessageClass                        66 LcnReadUnread          26 MessageId            26
                      104 UserConfigurationXmlStream         866 MessageClass           54 MessageId            26
                     1104 MessageClass                        56 MessageId              26 FolderId             26
                     1105 MessageClass                        56 MessageId              26 FolderId             26
                     1106 MessageClass                        56 MessageId              26 FolderId             26
                     1107 MessageClass                        66 MessageId              26 FolderId             26
                     1108 MessageId                           26 FolderId               26 LcnCurrent           26

        The GetTopColumnSizes query processor accepts one or more columns as arguments.  It produces additional columns to indicate the identity and size (in bytes) of the largest column values for each row. This example specifically excludes two large property blob columns from consideration by subtracting them.

        Use the Processors table function to see the list of all supported query processors and a brief description of each.

        .Example
        Get-StoreQuery -Database $db -Query "select * from ParsePropertyBlob($blob, 'Message')"
        $row = Get-StoreQuery -Database $db -Query "select * from Mailbox where MailboxNumber = 189"
        $blob = $row.PropertyBlob
        Get-StoreQuery -Database $db -Query "select * from ParsePropertyBlob($blob)" | ft -a PropertyTag,PropertyType,PropertyValue

        PropertyTag  PropertyType PropertyValue
        -----------  ------------ -------------
        0E23:Int32   Int32        138
        0E9B:Int32   Int32        522240
        3007:SysTime SysTime      2013-07-13 06:05:24.1703461
        3401:Int32   Int32        153034
        3402:Int32   Int32        1655
        3403:Int32   Int32        842
        3404:Int32   Int32        64
        3421:Binary  Binary       0x0100000083EFBDA4444ED048
        3422:Binary  Binary       0x010000008AB52EB95860D048
        3425:Binary  Binary       0x01000000B7AC1A84C34FD048

        Store query includes ParsePropertyBlob, a table function which can extract information from a binary property blob. The value of the blob is passed as a parameter to ParsePropertyBlob.

        .Example
        Get-Storequery -Database "db" - Query "select top 5 MessageDocumentId, DatabasePageNumber from Message where MailboxPartitionNumber = 125"

        MessageDocumentId                                          DatabasePageNumber
        -----------------                                          ------------------
                        1                                                      380888
                        2                                                      380888
                        3                                                      380888
                        4                                                      380888
                        5                                                      380888

       DatabasePageNumber is a virtual column that can be retrieved from any physical table.

       Use the VirtualColumns table function to see the list of all supported virtual columns and a brief description of each.
    #>

    [CmdletBinding(DefaultParameterSetName="Database")]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "Database", Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Database,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Query,

        [Parameter(Mandatory = $true, ParameterSetName = "ProcessId")]
        [ValidateNotNullOrEmpty()]
        [string]
        $Server,

        [Parameter(Mandatory = $true, ParameterSetName = "ProcessId")]
        [int]
        $ProcessId,

        [Parameter()]
        [Switch]
        $Unlimited
    )

    if ($MyInvocation.BoundParameters.ContainsKey("Database"))
    {
        [PSObject]$private:mdb = Get-MailboxDatabase $Database -Status

        if ($mdb -ne $null)
        {
            if ($mdb.Mounted -eq $null -or $mdb.Mounted -eq $false)
            {
                throw "Database $Database is dismounted or inaccessible"
            }

            [string]$private:mountedOnServer = $mdb.MountedOnServer
            [int]$private:workerProcessId = $mdb.WorkerProcessId
        }
        else
        {
            throw "Database $Database is not found"
        }
    }
    else
    {
        [string]$private:mountedOnServer = $Server
        [int]$private:workerProcessId = $ProcessId
    }

    if ($Unlimited)
    {
        [string]$private:content = (Get-ExchangeDiagnosticInfo -Server $mountedOnServer -Process $workerProcessId -Argument $Query -Unlimited).Result
    }
    else
    {
        [string]$private:content = (Get-ExchangeDiagnosticInfo -Server $mountedOnServer -Process $workerProcessId -Argument $Query).Result
    }

    if ([string]::IsNullOrEmpty($content) -eq $false)
    {
        [xml]$private:xml = $content
        if ((Get-Member -InputObject $xml.Diagnostics -Name 'ProcessLocator') -ne $null)
        {
            throw "Get-ExchangeDiagnosticInfo was unable to bind to the RPC endpoint of the store worker on $mountedOnServer with process id $workerProcessId"
        }

        if ($xml.Diagnostics.Components.ManagedStoreQueryHandler -ne $null)
        {
            Transform-XmlQuery -Results $xml
        }
        else
        {
            $content
        }
    }
    else
    {
        throw "Get-ExchangeDiagnosticInfo returned no results"
    }
}

function Transform-XmlQuery
{
    <#
        .Description
        Convert XML from Get-ExchangeDiagnosticInfo to rows of PSObject (internal)
        .Parameter Results
        XML from Get-ExchangeDiagnosticInfo
    #>

    param
    (
        [Parameter(Position = 0, Mandatory = $true)]
        [ValidateNotNull()]
        [xml]$Results
    )

    [Hashtable]$private:columnNames = @{}
    [Hashtable]$private:columnTypes = @{}
    [bool]$private:duplicateColumnWarning = $false
    $private:warnings = @()

    if (get-Member -InputObject $Results.Diagnostics.Components.ManagedStoreQueryHandler.Results -Name "Truncated")
    {
        [bool]$private:isTruncated = $false
        [void][bool]::TryParse($Results.Diagnostics.Components.ManagedStoreQueryHandler.Results.Truncated, [ref]$isTruncated)

        if ($isTruncated)
        {
            $private:warning = "Result set was truncated; select fewer columns or refine the query conditions"
            $private:warnings += $private:warning
            write-Warning $private:warning
        }
    }

    if (get-Member -InputObject $Results.Diagnostics.Components.ManagedStoreQueryHandler.Results -Name "Interrupted")
    {
        [bool]$private:isInterrupted = $false
        [void][bool]::TryParse($Results.Diagnostics.Components.ManagedStoreQueryHandler.Results.Interrupted, [ref]$isInterrupted)

        if ($isInterrupted)
        {
            $private:warning = "Query was interrupted for taking too much time or reading too many rows; you may override these internal limits with the Unlimited switch"
            $private:warnings += $private:warning
            write-Warning $private:warning
        }
    }

    foreach ($private:column in $Results.Diagnostics.Components.ManagedStoreQueryHandler.Results.Columns.Column)
    {
        if ($columnNames.ContainsValue($column.Name) -eq $false)
        {
            $columnNames.Add($column.Index, $column.Name)
            $columnTypes.Add($column.Index, $column.Type)
        }
        elseif ($duplicateColumnWarning -eq $false)
        {
            $private:warning = "Result set contained one or more duplicate columns"
            $private:warnings += $private:warning
            write-Warning $private:warning
            $duplicateColumnWarning = $true
        }
    }

    if ($Results.Diagnostics.Components.ManagedStoreQueryHandler.Results.Rows.Row -ne $null)
    {
        foreach ($private:row in $Results.Diagnostics.Components.ManagedStoreQueryHandler.Results.Rows.Row)
        {
            [PSObject]$private:displayRow = new-Object PSObject

            foreach ($private:value in $row.Value)
            {
                if ($columnNames.ContainsKey($value.Index))
                {
                    if ([string]::IsNullOrEmpty($value.InnerText) -eq $false)
                    {
                        [PSObject]$private:displayValue = Convert-RowValue $columnTypes[$value.Index] $value
                    }
                    else
                    {
                        [PSObject]$private:displayValue = $value.InnerText
                    }

                    add-Member -MemberType NoteProperty -Name $columnNames[$value.Index] -InputObject $displayRow -Value $displayValue
                }
            }

            $displayRow
        }
    }
    else
    {
        $displayRow = new-Object PSObject

        foreach ($private:column in $Results.Diagnostics.Components.ManagedStoreQueryHandler.Results.Columns.Column)
        {
            add-Member -MemberType NoteProperty -Name $column.Name -InputObject $displayRow -Value ([String]::Empty)
        }

        $displayRow
    }

    # Emit warnings again so they are not missed if the console buffer wraps
    foreach ($private:warning in $private:warnings)
    {
        write-Warning $private:warning
    }
}

function Format-Text
{
    <#
        .Description
        Display certain value columns using alternative text formats
        .Parameter Rows
        One or more PSObjects returned from Get-StoreQuery
        .Parameter Properties
        Comma-separated list of property/column names to reformat (default behavior is to apply reformatting to all columns)
    #>

    param
    (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNull()]
        [PSObject[]]$Rows,

        [Parameter(Position = 1, Mandatory = $false)]
        [string[]]$Properties
    )

    begin
    {
        [Collections.ArrayList]$list = new-Object Collections.ArrayList

        if ($MyInvocation.BoundParameters.ContainsKey("Properties") -and $Properties -ne $null -and $Properties.Length -gt 0)
        {
            foreach ($private:prop in $Properties)
            {
                [void]$list.Add($prop)
            }
        }
    }

    process
    {
        # Get-Member for NoteProperty returns properties in alphabetical order, not in original order.
        # To preserve original order, use the Force parameter to obtain the special psextended property,
        # which lists NoteProperty columns added to the object in order. These statements remove decorations
        # from the content of psextended then split it into an array of strings.

        [PSObject]$private:outputRow = new-Object PSObject
        [string]$private:psExtValue = ($Rows | Get-Member -MemberType MemberSet -Force | where { $_.Name -eq "psextended" }).Definition
        [string]$private:columnNames = $psExtValue.Replace("{", "").Replace("}", "")
        [string[]]$private:props = @($columnNames.Split(", ", [StringSplitOptions]::RemoveEmptyEntries) | where { $_ -ne "psextended" })

        if ($props.Length -gt 0)
        {
            foreach ($private:propName in $props)
            {
                if ($list.Count -eq 0 -or $list.Contains($propName))
                {
                    [Type]$type = $Rows.($propName).GetType()

                    if ($type.FullName -eq "System.Int16" -or $type.FullName -eq "System.Int32" -or $type.FullName -eq "System.Int64")
                    {
                        add-Member -MemberType NoteProperty -InputObject $outputRow -Name $propName -Value ("0x" + ($Rows.($propName)).ToString("X"))
                    }
                    elseif ($type.FullName -eq "System.DateTime")
                    {
                        add-Member -MemberType NoteProperty -InputObject $outputRow -Name $propName -Value (($Rows.($propName)).ToString("yyyy-MM-dd HH:mm:ss.fffffff"))
                    }
                    else
                    {
                        add-Member -MemberType NoteProperty -InputObject $outputRow -Name $propName -Value ($Rows.($propName))
                    }
                }
                else
                {
                    add-Member -MemberType NoteProperty -InputObject $outputRow -Name $propName -Value ($Rows.($propName))
                }
            }
        }

        $outputRow
    }
}

function Convert-RowValue
{
    <#
        .Description
        Convert a string value to a typed value (internal)
        .Parameter ValueType
        Type of the value
        .Parameter ValueText
        String representation of the value
    #>

    param
    (
        [Parameter(Position = 0, Mandatory = $true)]
        [ValidateNotNull()]
        [string]$ValueType,

        [Parameter(Position = 1, Mandatory = $true)]
        [ValidateNotNull()]
        [Xml.XmlElement]$ValueElement
    )

    [string]$private:valueText = $ValueElement.InnerText

    if ($valueText -eq "NULL")
    {
        return $valueText
    }
    else
    {
        switch ($ValueType)
        {
            "System.Boolean"
            {
                return $valueText.Equals("true", [StringComparison]::OrdinalIgnoreCase)
            }
            "System.Byte"
            {
                return [Byte]$valueText
            }
            "System.Char"
            {
                return [Char]$valueText
            }
            "System.Int16"
            {
                return [Int16]$valueText
            }
            "System.Int32"
            {
                return [Int32]$valueText
            }
            "System.Int64"
            {
                return [Int64]$valueText
            }
            "System.Single"
            {
                return [Single]$valueText
            }
            "System.Double"
            {
                return [Double]$valueText
            }
            "System.Guid"
            {
                return [Guid]$valueText
            }
            "System.DateTime"
            {
                return [DateTime]$valueText
            }
            "System.String"
            {
                return $valueText
            }
            "System.Byte[]"
            {
                return $valueText
            }
            "System.Int16[]"
            {
                [Int16[]]$array = @()

                foreach ($private:item in $ValueElement.Item)
                {
                    $array += [Int16]$item.InnerText
                }

                return $array
            }
            "System.Int32[]"
            {
                [Int32[]]$array = @()

                foreach ($private:item in $ValueElement.Item)
                {
                    $array += [Int32]$item.InnerText
                }

                return $array
            }
            "System.Int64[]"
            {
                [Int64[]]$array = @()

                foreach ($private:item in $ValueElement.Item)
                {
                    $array += [Int64]$item.InnerText
                }

                return $array
            }
            "System.String[]"
            {
                [string[]]$array = @()

                foreach ($private:item in $ValueElement.Item)
                {
                    $array += $item.InnerText
                }

                return $array
            }
            default
            {
                return $valueText
            }
        }
    }
}

function Test-StoreQuery
{
    <#
        .Description
        Check that the Get-ExchangeDiagnosticInfo on which Get-StoreQuery is built can communicate with store
        .Parameter Server
        Mailbox server to check.
        If no database is specified, the tool will attempt to connect to all databases on this server, both active and passive.
        .Parameter Database
        Mailbox database to check.
        If no server is specified, The tool will attempt to connect to the Active copy. Accepts wild cards.
    #>

    param(
        [Parameter()]
        [string]
        $Server = "",

        [Parameter()]
        [string]
        $Database = "")

    [PSObject[]]$private:mdbs = $null;
    if ([string]::IsNullOrEmpty($Database) -eq $false)
    {
        $mdbs = [PSObject[]](Get-MailboxDatabase -Identity $Database -Status)
    }
    elseif ([string]::IsNullOrEmpty($Server) -eq $false)
    {
        $mdbs = [PSObject[]](Get-MailboxDatabase -Server $Server -Status)
    }
    else
    {
        throw "Must specify Server or Database"
    }

    if ($mdbs -eq $null -or $mdbs.Length -eq 0)
    {
        throw "Failed to retrieve any databases"
    }

    [PSObject[]]$private:results = @()
    foreach ($private:mdb in $mdbs)
    {
        [string]$private:mdbServer = $mdb.MountedOnServer
        if ([string]::IsNullOrEmpty($Server) -eq $false)
        {
            $mdbServer = $Server
        }

        [PSObject]$private:result = new-Object PSObject
        add-Member -InputObject $result -MemberType NoteProperty -Name Server -Value $mdbServer
        add-Member -InputObject $result -MemberType NoteProperty -Name Database -Value $mdb.Name
        add-Member -InputObject $result -MemberType NoteProperty -Name Result -Value "Success"
        add-Member -InputObject $result -MemberType NoteProperty -Name Error -Value $null

        [string]$private:content = (Get-ExchangeDiagnosticInfo -Server $mdbServer -Process $mdb.WorkerProcessId -Argument 'select top 1 * from Mailbox').Result
        if ([string]::IsNullOrEmpty($content))
        {
            $result.Result = "Failed"
            $result.Error = "Get-ExchangeDiagnosticInfo failed to retrieve any content from the store worker"
        }
        else
        {
            [xml]$private:xml = $content
            if ((Get-Member -InputObject $xml.Diagnostics -Name 'ProcessLocator') -ne $null)
            {
                $result.Result = "Failed"
                $result.Error = "Get-ExchangeDiagnosticInfo was unable to bind to the RPC endpoint of the store worker"
            }
        }

        $results += $result
    }

    $results
}

function Get-StoreQueryCatalog
{
    param
    (
        [Parameter(Mandatory = $true)]
        $Database
    )

    $catalog = Get-StoreQuery -Database $Database -Query "SELECT * FROM Catalog"
    $catalog += Get-Command *StoreQuery* | % `
    {
        $function = New-Object PSObject | Select `
            TableName, `
            TableType, `
            PartitionKey, `
            Parameters, `
            Visibilty

        $function.TableName = $_.Name
        $function.TableType = "PowerShell Function"
        $function
    }

    $catalog | Sort TableName
}

function ForEach-StoreQueryMailbox
{
    <#
        .Description
        Execute a block of script for each mailbox found using Get-StoreQuery
        .Parameter Database
        Mailbox database to interrogate
        .Parameter MailboxFilter
        Optional filter to append to the where clause of the query of the mailbox table.
        .Parameter ExcludedMailboxStates
        Mailbox states to exclude (by default, CreatedByMove, DisabledMailbox, SoftDeletedMailbox,
        and MRSSoftDeletedMailbox are excluded).
        .Parameter ExcludedMailboxTypes
        Mailbox types to exclude (by default, none are excluded)
        .Parameter PerMailboxScript
        Optional script to execute against each mailbox table row. (If not specified, the row
        itself is returned.)
    #>

    param
    (
        [Parameter(Mandatory=$true)]
        $Database,

        [string]
        $MailboxFilter,

        [ValidateSet("CreatedByMove", "DisabledMailbox", "SoftDeletedMailbox", "MRSSoftDeletedMailbox")]
        [string[]]
        $ExcludedMailboxStates = @("CreatedByMove", "DisabledMailbox", "SoftDeletedMailbox", "MRSSoftDeletedMailbox"),

        [ValidateSet("PrimaryUserMailbox", "ArchiveUserMailbox", "SharedMailbox", "TeamMailbox", "PublicFolderPrimary", "PublicFolderSecondary")]
        [string[]]
        $ExcludedMailboxTypes = @(),

        [ScriptBlock]
        $PerMailboxScript
    )

    Write-Verbose "Excluded mailbox states: $ExcludedMailboxStates"
    Write-Verbose "Excluded mailbox types: $ExcludedMailboxTypes"

    $query = "SELECT *, MailboxMiscFlags, MailboxType, MailboxTypeDetail FROM Mailbox WHERE MailboxGuid != NULL"
    if (-not [String]::IsNullOrEmpty($MailboxFilter))
    {
        $query = "$query AND $MailboxFilter"
    }

    $dbs = @(Get-MailboxDatabase -Identity $Database)
    foreach ($db in $dbs)
    {
        $db = Get-MailboxDatabase $db -Status
        Write-Verbose "DB $($db.Name): MountedOnServer = $($db.MountedOnServer) WorkerProcessId = $($db.WorkerProcessId)"

        Write-Verbose "DB $($db.Name): $query"
        $mbs = @(Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query -Unlimited)

        foreach ($mb in $mbs)
        {
            if ([String]::IsNullOrEmpty($mb.MailboxNumber))
            {
                continue
            }

            $mailboxMiscFlags = $mb.p68060003
            $mailboxType = $mb.p3DBC0003
            $mailboxTypeDetail = $mb.p3DA60003

            # Exclude mailbox states we are not interested in, if requested.
            # p68060003 is MailboxMiscFlags.  Check for CreatedByMove, DisabledMailbox, SoftDeletedMailbox,
            # and MRSSoftDeletedMailbox.  All other flags are ignored here.
            if ($ExcludedMailboxStates -contains "CreatedByMove" -and ($mailboxMiscFlags -band 0x10) -ne 0)
            {
                Write-Verbose "DB $($db.Name): MB $($mb.MailboxNumber): Skipped (ExcludedMailboxState = CreatedByMove)"
                continue
            }

            if ($ExcludedMailboxStates -contains "DisabledMailbox" -and ($mailboxMiscFlags -band 0x40) -ne 0)
            {
                Write-Verbose "DB $($db.Name): MB $($mb.MailboxNumber): Skipped (ExcludedMailboxState = DisabledMailbox)"
                continue
            }

            if ($ExcludedMailboxStates -contains "SoftDeletedMailbox" -and ($mailboxMiscFlags -band 0x80) -ne 0)
            {
                Write-Verbose "DB $($db.Name): MB $($mb.MailboxNumber): Skipped (ExcludedMailboxState = SoftDeletedMailbox)"
                continue
            }

            if ($ExcludedMailboxStates -contains "MRSSoftDeletedMailbox" -and ($mailboxMiscFlags -band 0x100) -ne 0)
            {
                Write-Verbose "DB $($db.Name): MB $($mb.MailboxNumber): Skipped (ExcludedMailboxState = MRSSoftDeletedMailbox)"
                continue
            }

            # Exclude mailbox types we are not interested in, if requested.
            # p3DBC0003 is MailboxType. p3DA60003 is MailboxTypeDetail. p68060003 is MailboxMiscFlags.  If we
            # have a private user mailbox (MailboxType = 0 and MailboxTypeDetail = 1), check if this is a primary
            # or and archive and exclude as appropriate by looking at MailboxMiscFlags. For other mailbox types
            # (SharedMailbox, TeamMailbox, PublicFolderPrimary, or PublicFolderSecondary, we only look at
            # MailboxType and MailboxTypeDetail
            if ($mailboxType -eq 0 -and $mailboxTypeDetail -eq 1)
            {
                if ($ExcludedMailboxTypes -contains "PrimaryUserMailbox" -and ($mailboxMiscFlags -band 0x20) -eq 0)
                {
                    Write-Verbose "DB $($db.Name): MB $($mb.MailboxNumber): Skipped (ExcludedMailboxType = PrimaryUserMailbox)"
                    continue
                }

                if ($ExcludedMailboxTypes -contains "ArchiveUserMailbox" -and ($mailboxMiscFlags -band 0x20) -ne 0)
                {
                    Write-Verbose "DB $($db.Name): MB $($mb.MailboxNumber): Skipped (ExcludedMailboxType = ArchiveUserMailbox)"
                    continue
                }
            }

            if ($ExcludedMailboxTypes -contains "SharedMailbox" -and $mailboxType -eq 0 -and $mailboxTypeDetail -eq 2)
            {
                Write-Verbose "DB $($db.Name): MB $($mb.MailboxNumber): Skipped (ExcludedMailboxType = SharedMailbox)"
                continue
            }

            if ($ExcludedMailboxTypes -contains "TeamMailbox" -and $mailboxType -eq 0 -and $mailboxTypeDetail -eq 3)
            {
                Write-Verbose "DB $($db.Name): MB $($mb.MailboxNumber): Skipped (ExcludedMailboxType = TeamMailbox)"
                continue
            }

            if ($ExcludedMailboxTypes -contains "PublicFolderPrimary" -and $mailboxType -eq 1 -and $mailboxTypeDetail -eq 0)
            {
                Write-Verbose "DB $($db.Name): MB $($mb.MailboxNumber): Skipped (ExcludedMailboxType = PublicFolderPrimary)"
                continue
            }

            if ($ExcludedMailboxTypes -contains "PublicFolderSecondary" -and $mailboxType -eq 2 -and $mailboxTypeDetail -eq 0)
            {
                Write-Verbose "DB $($db.Name): MB $($mb.MailboxNumber): Skipped (ExcludedMailboxType = PublicFolderSecondary)"
                continue
            }

            Write-Verbose "DB $($db.Name): MB $($mb.MailboxNumber): TenantHint = $($mb.TenantHint) MailboxGuid = $($mb.MailboxGuid)"

            if ($PerMailboxScript -eq $null)
            {
                $mb
            }
            else
            {
                $PerMailboxScript.Invoke($db, $mb)
            }
        }
    }
}

function ForEach-StoreQueryFolder
{
    <#
        .Description
        Execute a block of script for each folder in each mailbox found using Get-StoreQuery
        .Parameter Database
        Mailbox database to interrogate
        .Parameter MailboxFilter
        Optional filter to append to the where clause of the query of the mailbox table.
        .Parameter ExcludedMailboxStates
        Mailbox states to exclude (by default, CreatedByMove, DisabledMailbox, SoftDeletedMailbox,
        and MRSSoftDeletedMailbox are excluded).
        .Parameter ExcludedMailboxTypes
        Mailbox types to exclude (by default, none are excluded)
        .Parameter MailboxFilter
        Optional filter to append to the where clause of the query of the folder table.
        .Parameter PerFolderScript
        Optional script to execute against each folder table row. (If not specified, the row
        itself is returned.)
    #>

    param
    (
        [Parameter(Mandatory=$true)]
        $Database,

        [string]
        $MailboxFilter,

        [ValidateSet("CreatedByMove", "DisabledMailbox", "SoftDeletedMailbox", "MRSSoftDeletedMailbox")]
        [string[]]
        $ExcludedMailboxStates = @("CreatedByMove", "DisabledMailbox", "SoftDeletedMailbox", "MRSSoftDeletedMailbox"),

        [ValidateSet("PrimaryUserMailbox", "ArchiveUserMailbox", "SharedMailbox", "TeamMailbox", "PublicFolderPrimary", "PublicFolderSecondary")]
        [string[]]
        $ExcludedMailboxTypes = @(),

        [string]
        $FolderFilter,

        [ScriptBlock]
        $PerFolderScript
    )

    ForEach-Mailbox -Database $Database -MailboxFilter $MailboxFilter -ExcludedMailboxStates $ExcludedMailboxStates -ExcludedMailboxTypes $ExcludedMailboxTypes -PerMailboxScript `
    {
        param($DatabaseObject, $MailboxObject)

        $query = "SELECT * FROM Folder WHERE MailboxNumber = $($MailboxObject.MailboxNumber)"
        if (-not [String]::IsNullOrEmpty($FolderFilter))
        {
            $query = "$query AND $FolderFilter"
        }

        Write-Verbose "DB $($DatabaseObject.Name): $query"
        $folders = @(Get-StoreQuery -Server $DatabaseObject.MountedOnServer -ProcessId $DatabaseObject.WorkerProcessId -Query $query)

        foreach ($folder in $folders)
        {
            if ([String]::IsNullOrEmpty($folder.MailboxNumber))
            {
                continue
            }

            Write-Verbose "DB $($DatabaseObject.Name): MB $($MailboxObject.MailboxNumber): FLDR $($folder.FolderId)"

            if ($PerFolderScript -eq $null)
            {
                $folder
            }
            else
            {
                $PerFolderScript.Invoke($DatabaseObject, $MailboxObject, $folder)
            }
        }
    }
}

function Get-StoreQueryRecipientTable
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $Database,

        [Parameter(Mandatory = $true, ParameterSetName = "MessageDocumentId")]
        [Parameter(Mandatory = $true, ParameterSetName = "FidIsHiddenMid")]
        [Parameter(Mandatory = $true, ParameterSetName = "EntryId")]
        [int]
        $MailboxPartitionNumber,

        [Parameter(Mandatory = $true, ParameterSetName = "MessageDocumentId")]
        [int]
        $MessageDocumentId,

        [Parameter(Mandatory = $true, ParameterSetName = "FidIsHiddenMid")]
        [string]
        $FolderId,

        [Parameter(Mandatory = $true, ParameterSetName = "FidIsHiddenMid")]
        [bool]
        $IsHidden,

        [Parameter(Mandatory = $true, ParameterSetName = "FidIsHiddenMid")]
        [string]
        $MessageId,

        [Parameter(Mandatory = $true, ParameterSetName = "EntryId")]
        [string]
        $EntryId,

        [Parameter(Mandatory = $true, ParameterSetName = "RecipientList")]
        [string]
        $RecipientList,

        [Parameter(Mandatory = $true, ParameterSetName = "MessageRow")]
        $MessageRow
    )

    $dbs = @(Get-MailboxDatabase -Identity $Database)
    if ($dbs.Count -eq 0)
    {
        Write-Error "Database $Database not found."
    }
    elseif ($dbs.Count -gt 1)
    {
        Write-Error "Database $Database matches more than one database."
    }
    else
    {
        $db = Get-MailboxDatabase $dbs[0].Name -Status

        # Get the recipient list blob based on the ParameterSet

        # MessageDocumentId ParameterSet
        if ($MyInvocation.BoundParameters.ContainsKey("MessageDocumentId"))
        {
            $query = "SELECT MessageDocumentId, RecipientList FROM Message WHERE MailboxPartitionNumber = $MailboxPartitionNumber AND MessageDocumentId = $MessageDocumentId"
            $message = Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query
            if ($message.MessageDocumentId -ne "")
            {        
                $recipientListBlob = $message.RecipientList
            }
            else
            {
                Write-Error "Message not found"
            }
        }

        # FidIsHiddenMid ParameterSet
        if ($MyInvocation.BoundParameters.ContainsKey("FolderId"))
        {
            $query = "SELECT MessageDocumentId, RecipientList FROM Message WHERE MailboxPartitionNumber = $MailboxPartitionNumber AND FolderId = $FolderId AND IsHidden = $IsHidden AND MessageId = $MessageId"
            $message = Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query
            if ($message.MessageDocumentId -ne "")
            {        
                $recipientListBlob = $message.RecipientList
            }
            else
            {
                Write-Error "Message not found"
            }
        }

        # EntryId ParameterSet
        if ($MyInvocation.BoundParameters.ContainsKey("EntryId"))
        {
            $query = "SELECT * FROM ParseEntryId($MailboxPartitionNumber, $EntryId)"
            $entryIdParts = Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query
            $recipientListBlob = "NULL"
            if ($entryIdParts.EidType -eq "eitLTPrivateMessage")
            {
                $query = "SELECT MessageDocumentId, RecipientList FROM Message WHERE MailboxPartitionNumber = $MailboxPartitionNumber AND FolderId = $($entryIdParts.FolderId) AND IsHidden = false AND MessageId = $($entryIdParts.MessageId)"
                $message = Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query
                if ($message.MessageDocumentId -ne "")
                {
                    $recipientListBlob = $message.RecipientList
                }
                else
                {
                    $query = "SELECT MessageDocumentId, RecipientList FROM Message WHERE MailboxPartitionNumber = $MailboxPartitionNumber AND FolderId = $($entryIdParts.FolderId) AND IsHidden = true AND MessageId = $($entryIdParts.MessageId)"
                    $message = Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query
                    if ($message.MessageDocumentId -ne "")
                    {
                        $recipientListBlob = $message.RecipientList
                    }
                    else
                    {
                        Write-Error "Message not found"
                    }
                }
            }
        }

        # RecipientList ParameterSet
        if ($MyInvocation.BoundParameters.ContainsKey("RecipientList"))
        {
            $recipientListBlob = $RecipientList
        }

        # MessageRow ParameterSet
        if ($MyInvocation.BoundParameters.ContainsKey("MessageRow"))
        {
            if (($MessageRow | Get-Member | ?{$_.Name -eq "RecipientList"} | Measure).Count -eq 1)
            {
                $recipientListBlob = $MessageRow.RecipientList
            }
            else
            {
                Write-Error "MessageRow provided does not contain RecipientList."
            }
        }

        if ($recipientListBlob -ne $null -and $recipientListBlob -ne "NULL")
        {
            # Get the individual recipient row blobs from the recipient list blob
            $query = "SELECT * FROM ParseMVBinaryBlob($recipientListBlob)"
            $recipientRowBlobs = @(Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query)

            $recipientCount = $recipientRowBlobs.Count
            $recipientRows = @{}
            $outputColumns = @("RowNumber")
            for ($i = 0; $i -lt $recipientCount; $i++)
            {
                # For each row, get the properties in the row
                # There are no named properties on the recipient table so there is no need to pass in a mailbox number in order to resolve them.
                $query = "SELECT PropertyId, PropertyValue FROM ParsePropertyBlob($($recipientRowBlobs[$i].Value), 'Recipient')"
                $recipientRows[$i] = @(Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query)

                # Construct the list of output properties (the union of all properties found on all rows)
                foreach ($property in $recipientRows[$i])
                {
                    # If we were unable to resolve a property against the message table schema then it will be output as 2345:Unicode
                    # Prefix the property tag with a p and replace : with _ so the output object can be manipulated in PowerShell
                    if ($property.PropertyId -match '^[0-9]')
                    {
                        $property.PropertyId = "p$($property.PropertyId)"
                    }

                    $property.PropertyId = $property.PropertyId.Replace(':', '_')
                    if ($outputColumns -notcontains $property.PropertyId)
                    {
                        $outputColumns += $property.PropertyId
                    }
                }
            }

            for ($i = 0; $i -lt $recipientCount; $i++)
            {
                $outputRow = New-Object PSObject | Select $outputColumns
                $outputRow.RowNumber = $i
                foreach ($property in $recipientRows[$i])
                {
                    Add-Member -MemberType NoteProperty -Name $property.PropertyId -InputObject $outputRow -Value $property.PropertyValue -Force
                }

                Write-Output $outputRow
            }
        }
    }
}

function Get-StoreQueryFolderViews
{
    [CmdletBinding(DefaultParameterSetName="Fid")]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $Database,

        [Parameter(Mandatory = $true)]
        [int]
        $MailboxNumber,

        [Parameter(ParameterSetName = "Fid")]
        [string]
        $FolderId,

        [Parameter(ParameterSetName = "EntryId")]
        [string]
        $EntryId,

        [Parameter(ParameterSetName = "LogicalIndexNumbers")]
        [int[]]
        $LogicalIndexNumbers
    )

    $dbs = @(Get-MailboxDatabase -Identity $Database)
    if ($dbs.Count -eq 0)
    {
        Write-Error "Database $Database not found."
    }
    elseif ($dbs.Count -gt 1)
    {
        Write-Error "Database $Database matches more than one database."
    }
    else
    {
        $db = Get-MailboxDatabase $dbs[0].Name -Status

        # Get the folder id based on the ParameterSet

        # Fid ParameterSet - nothing to do FolderId is passed in

        # EntryId ParameterSet
        if ($MyInvocation.BoundParameters.ContainsKey("EntryId"))
        {
            $query = "SELECT * FROM ParseEntryId($MailboxNumber, $EntryId)"
            $entryIdParts = Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query
            if ($entryIdParts.EidType -eq "eitLTPrivateFolder")
            {
                $FolderId = $entryIdParts.FolderId
            }
        }

        $query = "SELECT * FROM PseudoIndexControl WHERE MailboxNumber = $MailboxNumber"
        if (-not [String]::IsNullOrEmpty($FolderId))
        {
            $query += " AND FolderId = $FolderId"
        }

        if ($LogicalIndexNumbers -ne $null)
        {
            $query += " AND ("
            $first = $true
            foreach ($logicalIndexNumber in $LogicalIndexNumbers)
            {
                if (-not $first)
                {
                    $query += " OR "
                }

                $query += "LogicalIndexNumber = $logicalIndexNumber"
                $first = $false
            }

            $query += ")"
        }

        $indexControlEntries = Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query -Unlimited

        $indexDefinitionBlobHeaders = @{}
        $indexDefinitionColumns = @{}

        foreach ($indexControlEntry in $indexControlEntries)
        {
            $view = $indexControlEntry | Select `
                MailboxNumber, `
                FolderId, `
                PhysicalIndexNumber, `
                LogicalIndexNumber, `
                LogicalIndexVersion, `
                IndexSignature, `
                TableName, `
                IndexType, `
                FirstUpdateRecord, `
                LastReferenceDate, `
                PhysicalIndexKeyColumnCount, `
                LogicalIndexKeyColumnCount, `
                Culture, `
                Columns, `
                ViewColumns, `
                Condition, `
                IdentityColumnIndex, `
                BaseMessageViewLogicalIndexNumber, `
                BaseMessageViewInReverseOrder, `
                CategoryCount, `
                CategoryHeaderSortOverrides

            switch ($view.IndexType)
            {
                0 { $view.IndexType = "Messages" }
                1 { $view.IndexType = "Conversations" }
                2 { $view.IndexType = "SearchFolderBaseView" }
                3 { $view.IndexType = "SearchFolderMessages" }
                4 { $view.IndexType = "CategoryHeaders" }
                5 { $view.IndexType = "ConversationDeleteHistory" }
            }

            if (-not $indexDefinitionBlobHeaders.ContainsKey($indexControlEntry.PhysicalIndexNumber) `
                -or -not $indexDefinitionColumns.ContainsKey($indexControlEntry.PhysicalIndexNumber))
            {
                $query = "SELECT * FROM PseudoIndexDefinition WHERE PhysicalIndexNumber = $($indexControlEntry.PhysicalIndexNumber)"
                $indexDefinitionEntry = Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query

                $query = "SELECT * FROM IndexDefinitionBlobHeader($($indexDefinitionEntry.ColumnBlob))"
                $indexDefinitionBlobHeaders[$indexDefinitionEntry.PhysicalIndexNumber] = Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query

                $query = "SELECT * FROM IndexDefinitionBlob($($indexDefinitionEntry.ColumnBlob))"
                $indexDefinitionColumns[$indexDefinitionEntry.PhysicalIndexNumber] = Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query
            }

            $view.PhysicalIndexKeyColumnCount = $indexDefinitionBlobHeaders[$indexControlEntry.PhysicalIndexNumber].keyColumnCount
            $view.Culture = [CultureInfo]$indexDefinitionBlobHeaders[$indexControlEntry.PhysicalIndexNumber].lcid
            $view.IdentityColumnIndex = $indexDefinitionBlobHeaders[$indexControlEntry.PhysicalIndexNumber].identityColumnIndex

            $query = "SELECT * FROM ColumnMappingBlobHeader($($indexControlEntry.ColumnMappings))"
            $columnMappingBlobHeader = Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query

            $view.LogicalIndexKeyColumnCount = $columnMappingBlobHeader.keyColumnCount

            $query = "SELECT * FROM ColumnMappingBlob($($indexControlEntry.ColumnMappings))"
            $logicalIndexColumns = Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query

            $view.Columns = @()
            for ($i = 0; $i -lt $indexDefinitionColumns[$indexControlEntry.PhysicalIndexNumber].Count; $i++)
            {
                $physicalIndexColumnIndex = $i
                $logicalIndexColumnIndex = $i - 2
                $column = New-Object PSObject | Select `
                    Property, `
                    Type, `
                    LogicalSort, `
                    PhysicalSort, `
                    FixedLength, `
                    MaximumLength,
                    IsLogicalIndexKeyColumn, `
                    IsPhysicalIndexKeyColumn, `
                    PhysicalIndexColumnName

                if ($logicalIndexColumnIndex -ge 0 `
                    -and $logicalIndexColumnIndex -lt $logicalIndexColumns.Count)
                {
                    if (-not [String]::IsNullOrEmpty($logicalIndexColumns[$logicalIndexColumnIndex].propName))
                    {
                        $column.Property += $logicalIndexColumns[$logicalIndexColumnIndex].propName
                    }
                    else
                    {
                        $column.Property += "0x" + $logicalIndexColumns[$logicalIndexColumnIndex].propId.ToString("X8")
                    }
                }

                switch ($indexDefinitionColumns[$indexControlEntry.PhysicalIndexNumber][$physicalIndexColumnIndex].columnType)
                {
                    -2147483617 { $column.Type = "Unicode" } # Unicode with UseLinguisticCasingRulesFlag set
                    0 { $column.Type = "Unspecified" }
                    1 { $column.Type = "Null" }
                    2 { $column.Type = "Int16" }
                    3 { $column.Type = "Int32" }
                    4 { $column.Type = "Real32" }
                    5 { $column.Type = "Real64" }
                    6 { $column.Type = "Currency" }
                    7 { $column.Type = "AppTime" }
                    10 { $column.Type = "Error" }
                    11 { $column.Type = "Boolean" }
                    13 { $column.Type = "Object" }
                    20 { $column.Type = "Int64" }
                    30 { $column.Type = "String8" }
                    31 { $column.Type = "Unicode" }
                    64 { $column.Type = "SysTime" }
                    72 { $column.Type = "Guid" }
                    251 { $column.Type = "SvrEid" }
                    253 { $column.Type = "SRestriction" }
                    254 { $column.Type = "Actions" }
                    258 { $column.Type = "Binary" }
                    4095 { $column.Type = "Invalid" }
                    4096 { $column.Type = "MVFlag" }
                    4097 { $column.Type = "MVNull" }
                    4098 { $column.Type = "MVInt16" }
                    4099 { $column.Type = "MVInt32" }
                    4100 { $column.Type = "MVReal32" }
                    4101 { $column.Type = "MVReal64" }
                    4102 { $column.Type = "MVCurrency" }
                    4103 { $column.Type = "MVAppTime" }
                    4116 { $column.Type = "MVInt64" }
                    4126 { $column.Type = "MVString8" }
                    4127 { $column.Type = "MVUnicode" }
                    4160 { $column.Type = "MVSysTime" }
                    4168 { $column.Type = "MVGuid" }
                    4354 { $column.Type = "MVBinary" }
                    8191 { $column.Type = "MVInvalid" }
                    8192 { $column.Type = "MVInstance" }
                    12289 { $column.Type = "MVINull" }
                    12290 { $column.Type = "MVIInt16" }
                    12291 { $column.Type = "MVIInt32" }
                    12292 { $column.Type = "MVIReal32" }
                    12293 { $column.Type = "MVIReal64" }
                    12294 { $column.Type = "MVICurrency" }
                    12295 { $column.Type = "MVIAppTime" }
                    12308 { $column.Type = "MVIInt64" }
                    12318 { $column.Type = "MVIString8" }
                    12319 { $column.Type = "MVIUnicode" }
                    12352 { $column.Type = "MVISysTime" }
                    12360 { $column.Type = "MVIGuid" }
                    12546 { $column.Type = "MVIBinary" }
                    16383 { $column.Type = "MVIInvalid" }
                }

                $column.FixedLength = $indexDefinitionColumns[$indexControlEntry.PhysicalIndexNumber][$physicalIndexColumnIndex].fixedLength
                $column.MaximumLength = $indexDefinitionColumns[$indexControlEntry.PhysicalIndexNumber][$physicalIndexColumnIndex].maxLength

                if ($logicalIndexColumnIndex -ge 0)
                {
                    $column.IsLogicalIndexKeyColumn = $logicalIndexColumnIndex -lt $view.LogicalIndexKeyColumnCount
                    if ($column.IsLogicalIndexKeyColumn)
                    {
                        if ($indexDefinitionColumns[$indexControlEntry.PhysicalIndexNumber][$physicalIndexColumnIndex].ascending)
                        {
                            $column.LogicalSort = "Ascending"
                        }
                        else
                        {
                            $column.LogicalSort = "Descending"
                        }
                    }
                }

                $column.IsPhysicalIndexKeyColumn = $physicalIndexColumnIndex -lt $view.PhysicalIndexKeyColumnCount
                if ($column.IsPhysicalIndexKeyColumn)
                {
                    if ($indexDefinitionColumns[$indexControlEntry.PhysicalIndexNumber][$physicalIndexColumnIndex].ascending)
                    {
                        $column.PhysicalSort = "Ascending"
                    }
                    else
                    {
                        $column.PhysicalSort = "Descending"
                    }
                }

                if ($physicalIndexColumnIndex -eq 0)
                {
                    $column.PhysicalIndexColumnName = "MailboxNumber"
                }
                elseif ($physicalIndexColumnIndex -eq 1)
                {
                    $column.PhysicalIndexColumnName = "LogicalIndexNumber"
                }
                else
                {
                    $column.PhysicalIndexColumnName = "C$($logicalIndexColumnIndex + 1)"
                }

                $view.Columns += $column
            }

            for ($i = 2; $i -lt $view.Columns.Count; $i++)
            {
                if (-not [String]::IsNullOrEmpty($view.ViewColumns))
                {
                    $view.ViewColumns += ", "
                }

                if ($view.Columns[$i].IsLogicalIndexKeyColumn)
                {
                    $sort = $view.Columns[$i].LogicalSort
                }
                else
                {
                    if ($view.Columns[$i].IsPhysicalIndexKeyColumn)
                    {
                        $sort = [String]::Format("Covering (physically {0})", $view.Columns[$i].PhysicalSort)
                    }
                    else
                    {
                        $sort = "Covering"
                    }
                }

                $view.ViewColumns += [String]::Format("{0} - {1}", $view.Columns[$i].Property, $sort)
            }

            if ($indexControlEntry.ConditionalIndex -ne "NULL")
            {
                $query = "SELECT * FROM ConditionalIndexMappingBlob($($indexControlEntry.ConditionalIndex))"
                $conditionalIndex = Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query

                $view.Condition = [String]::Format("{0} = {1}", $conditionalIndex.columnName, $conditionalIndex.columnValue)
            }

            if ($indexControlEntry.CategorizationInfo -ne "NULL")
            {
                $query = "SELECT * FROM ParseCategorizationInfo($($indexControlEntry.CategorizationInfo), $MailboxNumber)"
                $categorizationInfo = Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query

                $view.BaseMessageViewLogicalIndexNumber = $categorizationInfo.BaseMessageViewLogicalIndexNumber
                $view.BaseMessageViewInReverseOrder = $categorizationInfo.BaseMessageViewInReverseOrder
                $view.CategoryCount = $categorizationInfo.CategoryCount
                $view.CategoryHeaderSortOverrides  = $categorizationInfo.CategoryHeaderSortOverrides
            }

            Write-Output $view
        }
    }
}

# Enumerates the search folders in a specified mailbox
# and dumps information for each in a readable form.
function Get-StoreQuerySearchFolders
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $Database,

        [Parameter(Mandatory = $true)]
        [int]
        $MailboxNumber,

        [Parameter(Mandatory = $false)]
        [string]
        $FolderId
    )

    $dbs = @(Get-MailboxDatabase -Identity $Database)
    if ($dbs.Count -eq 0)
    {
        Write-Error "Database $Database not found."
    }
    elseif ($dbs.Count -gt 1)
    {
        Write-Error "Database $Database matches more than one database."
    }
    else
    {
        # Now that we've verified that the specified paramater maps to
        # exactly one database, use -Status to determine where the database
        # is currently mounted as well as the worker process id. We do this
        # as a small perf optimisation because we're going to be calling
        # Get-StoreQuery multiple times, so it's more efficient to pass
        # those params directly to Get-StoreQuery instead of having
        # Get-StoreQuery compute the mapping every time).
        $db = Get-MailboxDatabase -Identity $dbs[0].Name -Status


        # Retrieve the folder id of the MaterializedRestrictions special
        # folder so that we can identify materialized restriction search
        # folders. The MaterializedRestriction special folder has a
        # SpecialFolderNumber of 21, but there is no index on
        # SpecialFolderNumber, so to avoid a sequential scan of the Folder
        # table, we use the fact that the MaterializedRestrictions special
        # folder has no parent folder. So we can leverage the index on
        # ParentFolderId to scan for just the folders with no parent (there
        # shoud only be a few).
        $query = "SELECT FolderId,SpecialFolderNumber FROM Folder WHERE MailboxNumber = $MailboxNumber AND ParentFolderId = '0x0000000000000000000000000000000000000000000000000000'"
        $foldersWithNoParent = @(Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId  -Query $query -Unlimited)

        [string]$mailboxRootFolderId = [String]::Empty
        [string]$restrictionsFolderId = [String]::Empty
        foreach ($folderWithNoParent in $foldersWithNoParent)
        {
            if ($folderWithNoParent.SpecialFolderNumber -eq 21)
            {
                # Found the MaterializedRestrictions special folder.
                $restrictionsFolderId = $folderWithNoParent.FolderId
            }
            elseif ($folderWithNoParent.SpecialFolderNumber -eq 1)
            {
                # Found the mailbox root folder, which we'll need below to
                # help find the Finder special folder.
                $mailboxRootFolderId = $folderWithNoParent.FolderId
            }
        }

        # Retrieve the folder id of the Finder special folder so we can
        # identify search folders under that folder. Search folders under
        # Finder are of particular interest because that's where search
        # folders created by Outlook and OWA are typically located and also
        # because we permit search folders under Finder to be aged-out even if
        # no age-out-related properties are specifically set on the search
        # folder. The Finder special folder is located under the mailbox root
        # folder and has a hard-coded SpecialFolderNumber of 2.
        [string]$finderFolderId = [String]::Empty
        if (-not [String]::IsNullOrEmpty($mailboxRootFolderId))
        {
            $query = "SELECT FolderId,SpecialFolderNumber FROM Folder WHERE MailboxNumber = $MailboxNumber AND ParentFolderId = $mailboxRootFolderId AND SpecialFolderNumber = 2"
            $finderFolder = @(Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId  -Query $query -Unlimited)
            if ($finderFolder.Count -eq 1)
            {
                $finderFolderId = $finderFolder[0].FolderId
            }
        }
        
        # Indicate whether or not we were successful in determining the
        # MaterializedRestrictions special folder.
        if ([String]::IsNullOrEmpty($restrictionsFolderId))
        {
            Write-Warning "Unable to determine MaterializedRestrictions folder. The IsMaterializedRestriction property of each search folder in the result set will not be accurate."
        }
        else
        {
            Write-Verbose "Using MaterializedRestrictions FolderId=$restrictionsFolderId."
        }

        # Indicate whether or not we were successful in determining the
        # Finder special folder.
        if ([String]::IsNullOrEmpty($finderFolderId))
        {
            Write-Warning "Unable to determine Finder folder. The IsFinder property of each search folder in the result set will not be accurate."
        }
        else
        {
            Write-Verbose "Using Finder FolderId=$finderFolderId."
        }

        # Now fetch the list of search folders for the mailbox, including
        # AllowAgeOut (p361F000B) and SearchFolderAgeOutTimeout (p36470003).
        # NOTE: We have to break up p361F000B into two substrings in order
        # to work around FxCop security scan, which seems to think that the
        # full string resembles a password.
        [string]$propertiesToFetch =
            "DisplayName," + `
            "FolderId," + `
            "ParentFolderId," + `
            "LogicalIndexNumber," + `
            "CreationTime," + `
            "MessageCount," + `
            "HiddenItemCount," + `
            "FolderCount," + `
            "QueryCriteria," + `
            "ScopeFolders," + `
            "SetSearchCriteriaFlags," + `
            "SearchState," + `
            "p361F" + "000B," + `
            "p36470003"

        $additionalCriteria = ""
        if ($MyInvocation.BoundParameters.ContainsKey("FolderId"))
        {
            $additionalCriteria = " AND FolderId = $FolderId"
        }

        $query = "SELECT " + $propertiesToFetch + " FROM Folder WHERE MailboxNumber = $MailboxNumber AND QueryCriteria != null" + $additionalCriteria

        Write-Verbose "Retrieving search folders for mailbox $MailboxNumber."
        $searchFolders = @(Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query -Unlimited)

        Write-Warning "Parsing results for $($searchFolders.Count) search folders, which requires calling Get-StoreQuery up to 4 times per search folder, so this may take a very long time."

        [bool]$checkedRedaction = $false
        [bool]$valuesRedacted = $false

        # The following variables will be used to save off the last retrieved
        # value for QueryCriteria. We do this as an optimisation to save on
        # repeated parsing calls to Get-StoreQuery for mailboxes with a large
        # number of search folders having the same QueryCriteria.
        $previousQueryCriteria = [String]::Empty
        $previousFriendlyRestriction = [String]::Empty

        # We will be maintaining a dictionary to map the serialized scope
        # folder list to folder ids. We do this as an optimisation to save on
        # repeated parsing calls to Get-StoreQuery for mailboxes with a large
        # number of searches scoping the same folder. Currently, we only do
        # this for scope folder lists with a length of 1. This should cover
        # over 95% of search folders in the typical mailbox. I'm hesitant to
        # include scope folder lists of any length because we will be using
        # the list as the dictionary key and I'm uncertain of the implications
        # of keys of arbitrarily-long length.
        $serializedScopeFolderListToIdMap = @{}

        # Massage some of the column values to make then more readable.
        [int]$numSearchFoldersProcessed = 0
        foreach ($searchFolder in $searchFolders)
        {
            # Determine if redaction is in effect.
            if (!$checkedRedaction)
            {
                # Should be sufficient just to check the first entry, since
                # redaction affects either all entries or no entries.
                $valuesRedacted = __IsPropertyValueRedacted($searchFolder.DisplayName)
                $checkedRedaction = $true
            }

            # Parse the restriction.
            if ($searchFolder.QueryCriteria -eq $previousQueryCriteria)
            {
                # Current restriction matches previous restriction,
                # so no need to parse it.
                $friendlyRestriction = $previousFriendlyRestriction
            }
            else
            {
                $query = "SELECT * FROM ParseRestriction($($searchFolder.QueryCriteria), $MailboxNumber)"
                $parsedRestriction = Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query -Unlimited
                if ($parsedRestriction.DiagnosticQueryParserException -ne $null -and $valuesRedacted)
                {
                    # If a parsing error was encountered and redaction is in effect,
                    # assume the error was due to redaction.
                    $friendlyRestriction = "<REDACTED>"
                }
                else
                {
                    $friendlyRestriction = [string]$parsedRestriction.Value
                }

                # Save off the current restriction in case it's a
                # match with the next entry.
                $previousQueryCriteria = $searchFolder.QueryCriteria
                $previousFriendlyRestriction = $friendlyRestriction
            }

            # See if we can quickly determine the friendly version of the
            # serialized scope folder list (either because it's NULL or
            # because it's already in the dictionary) or if we have to parse
            # it.
            if ([string]$searchFolder.ScopeFolders -eq "NULL")
            {
                $friendlyScopeFolders = "NULL"
            }
            elseif ($searchFolder.ScopeFolders.Length -eq 26 -and $serializedScopeFolderListToIdMap.ContainsKey($searchFolder.ScopeFolders))
            {
                # The dictionary is only valid for scope folder lists with 1
                # entry (see above for an explanation why), so if the scope
                # folder list has only 1 entry and we found the serialized
                # scope folder list in the dictionary, then just retrieve the
                # friendly version from the dictionary. Note that we can tell
                # the scope folder list only has 1 entry in it by looking at
                # the length of the serialized scope folder list. A scope
                # folder list with only 1 entry will be serialized as a
                # 26-character string (2 characters for "0x", 8 characters for
                # the 4-byte number of entries in little-endian form, 4
                # characters for the 2-byte replid in little-endian form, and
                # 12 characters for the 6-byte globcnt in big-endian form).
                $friendlyScopeFolders = $serializedScopeFolderListToIdMap[$searchFolder.ScopeFolders]
            }
            else
            {
                # Parse the serialized scope folder list.
                $query = "SELECT * FROM ParseExchangeIdList($($searchFolder.ScopeFolders), $MailboxNumber)"
                $scopeFolderList = @(Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query -Unlimited)

                if ($scopeFolderList.Count -eq 0)
                {
                    $friendlyScopeFolders = "<Parse Error>"
                }
                elseif ($scopeFolderList.Count -eq 1)
                {
                    # Only one entry in the scope folder list, so report it as
                    # string for easier manipulation for operations such as
                    # grouping.
                    $friendlyScopeFolders = [string]$scopeFolderList[0].Value

                    # Try to append the DisplayName of the scope folder to
                    # facilitate debugging.
                    $query = "SELECT DisplayName FROM Folder WHERE MailboxNumber = $MailboxNumber AND FolderId = $friendlyScopeFolders"
                    $scopeFolders = @(Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query -Unlimited)
                    if ($scopeFolders.Count -eq 1 -and -not [String]::IsNullOrEmpty($scopeFolders[0].DisplayName))
                    {
                        $friendlyScopeFolders += " (" + $scopeFolders[0].DisplayName + ")"
                    }

                    # Save off the current ScopeFolders in case multiple searches
                    # scope the same folder. However, we currently only do this
                    # for scope folder lists with 1 entry in it (see explanation
                    # above for the reason for this limitation).
                    $serializedScopeFolderListToIdMap.Add($searchFolder.ScopeFolders, $friendlyScopeFolders)
                }
                else
                {
                    # Multiple entries in the scope folder list, so report it
                    # as an array of strings.
                    for ([int]$i = 0; $i -lt $scopeFolderList.Count; $i++)
                    {
                        $scopeFolderList[$i] = [string]$scopeFolderList[$i].Value
                    }
                    $friendlyScopeFolders = $scopeFolderList
                }
            }

            # Parse SetSearchCriteriaFlags, taking care to handle
            # missing SetSearchCriteriaFlags (which may occur if
            # SetSearchCriteria() never got called on the search
            # folder).
            if ([string]$searchFolder.SetSearchCriteriaFlags -eq "NULL")
            {
                $friendlySetSearchCriteriaFlags = "NULL"
            }
            else
            {
                $friendlySetSearchCriteriaFlags = [String]::Empty
                if (($searchFolder.SetSearchCriteriaFlags -band 0x00000001) -ne 0)
                {
                    $friendlySetSearchCriteriaFlags += ",Stop"
                }
                if (($searchFolder.SetSearchCriteriaFlags -band 0x00000002) -ne 0)
                {
                    $friendlySetSearchCriteriaFlags += ",Restart"
                }
                if (($searchFolder.SetSearchCriteriaFlags -band 0x00000004) -ne 0)
                {
                    $friendlySetSearchCriteriaFlags += ",Recursive"
                }
                if (($searchFolder.SetSearchCriteriaFlags -band 0x00000008) -ne 0)
                {
                    $friendlySetSearchCriteriaFlags += ",Shallow"
                }
                if (($searchFolder.SetSearchCriteriaFlags -band 0x00000010) -ne 0)
                {
                    $friendlySetSearchCriteriaFlags += ",Foreground"
                }
                if (($searchFolder.SetSearchCriteriaFlags -band 0x00000020) -ne 0)
                {
                    $friendlySetSearchCriteriaFlags += ",Background"
                }
                if (($searchFolder.SetSearchCriteriaFlags -band 0x00004000) -ne 0)
                {
                    $friendlySetSearchCriteriaFlags += ",UseCIForComplexQueries"
                }
                if (($searchFolder.SetSearchCriteriaFlags -band 0x00010000) -ne 0)
                {
                    $friendlySetSearchCriteriaFlags += ",ContentIndexed"
                }
                if (($searchFolder.SetSearchCriteriaFlags -band 0x00020000) -ne 0)
                {
                    $friendlySetSearchCriteriaFlags += ",NonContentIndexed"
                }
                if (($searchFolder.SetSearchCriteriaFlags -band 0x00040000) -ne 0)
                {
                    $friendlySetSearchCriteriaFlags += ",Static"
                }
                if (($searchFolder.SetSearchCriteriaFlags -band 0x00800000) -ne 0)
                {
                    $friendlySetSearchCriteriaFlags += ",FailOnForeignEID"
                }
                if (($searchFolder.SetSearchCriteriaFlags -band 0x01000000) -ne 0)
                {
                    $friendlySetSearchCriteriaFlags += ",StatisticsOnly"
                }
                if (($searchFolder.SetSearchCriteriaFlags -band 0x02000000) -ne 0)
                {
                    $friendlySetSearchCriteriaFlags += ",FailNonContentIndexedSearch"
                }
                if (($searchFolder.SetSearchCriteriaFlags -band 0x04000000) -ne 0)
                {
                    $friendlySetSearchCriteriaFlags += ",EstimateCountOnly"
                }

                # Report both the full bitfield as well as the friendly enum values.
                if ([String]::IsNullOrEmpty($friendlySetSearchCriteriaFlags))
                {
                    $friendlySetSearchCriteriaFlags = [String]::Format("0x{0:X} (None)", $searchFolder.SetSearchCriteriaFlags)
                }
                else
                {
                    # Strip off the leading comma from the list of friendly enum values.
                    $friendlySetSearchCriteriaFlags = [String]::Format("0x{0:X} ({1})", $searchFolder.SetSearchCriteriaFlags, $friendlySetSearchCriteriaFlags.Substring(1))
                }
            }

            # Parse SearchState, taking care to handlie missing
            # SearchState which may occur if SetSearchCriteria()
            # never got called on the search folder).
            $isInstantSearch = $false
            if ([string]$searchFolder.SearchState -eq "NULL")
            {
                $friendlySearchState = "NULL"
            }
            else
            {
                $friendlySearchState = [String]::Empty
                if (($searchFolder.SearchState -band 0x00000001) -ne 0)
                {
                    $friendlySearchState += ",Running"
                }
                if (($searchFolder.SearchState -band 0x00000002) -ne 0)
                {
                    $friendlySearchState += ",Rebuild"
                }
                if (($searchFolder.SearchState -band 0x00000004) -ne 0)
                {
                    $friendlySearchState += ",Recursive"
                }
                if (($searchFolder.SearchState -band 0x00000010) -ne 0)
                {
                    $friendlySearchState += ",Foreground"
                }
                if (($searchFolder.SearchState -band 0x00001000) -ne 0)
                {
                    $friendlySearchState += ",AccurateResults"
                }
                if (($searchFolder.SearchState -band 0x00002000) -ne 0)
                {
                    $friendlySearchState += ",PotentiallyInaccurateResults"
                }
                if (($searchFolder.SearchState -band 0x00010000) -ne 0)
                {
                    $friendlySearchState += ",Static"
                }
                if (($searchFolder.SearchState -band 0x00020000) -ne 0)
                {
                    $isInstantSearch = $true
                    $friendlySearchState += ",InstantSearch"
                }
                if (($searchFolder.SearchState -band 0x00080000) -ne 0)
                {
                    $friendlySearchState += ",StatisticsOnly"
                }
                if (($searchFolder.SearchState -band 0x00100000) -ne 0)
                {
                    $friendlySearchState += ",CiOnly"
                }
                if (($searchFolder.SearchState -band 0x00200000) -ne 0)
                {
                    $friendlySearchState += ",FullTextIndexQueryFailed"
                }
                if (($searchFolder.SearchState -band 0x00400000) -ne 0)
                {
                    $friendlySearchState += ",EstimateCountOnly"
                }
                if (($searchFolder.SearchState -band 0x01000000) -ne 0)
                {
                    $friendlySearchState += ",CiTotally"
                }
                if (($searchFolder.SearchState -band 0x02000000) -ne 0)
                {
                    $friendlySearchState += ",CiWithTwirResidual"
                }
                if (($searchFolder.SearchState -band 0x04000000) -ne 0)
                {
                    $friendlySearchState += ",TwirMostly"
                }
                if (($searchFolder.SearchState -band 0x08000000) -ne 0)
                {
                    $friendlySearchState += ",TwirTotally"
                }
                if (($searchFolder.SearchState -band 0x10000000) -ne 0)
                {
                    $friendlySearchState += ",Error"
                }

                # Report both the full bitfield as well as the friendly enum values.
                if ([String]::IsNullOrEmpty($friendlySearchState))
                {
                    $friendlySearchState = [String]::Format("0x{0:X} (None)", $searchFolder.SearchState)
                }
                else
                {
                    # Strip off the leading comma from the list of friendly enum values.
                    $friendlySearchState = [String]::Format("0x{0:X} ({1})", $searchFolder.SearchState, $friendlySearchState.Substring(1))
                }
            }            

            # Determine if the AllowAgeOut property was set.
            [bool]$allowAgeout = $false
            if ([string]$searchFolder.p361F000B -ne "NULL")
            {
                $allowAgeout = [bool]$searchFolder.p361F000B
            }

            # Determine if the SearchFolderAgeOutTimeout property was set.
            [int]$ageoutTimeout = 0
            if ([string]$searchFolder.p36470003 -ne "NULL")
            {
                $ageoutTimeout = [int]$searchFolder.p36470003
            }

            # Determine the number of views over the search folder.
            $query = "SELECT LastReferenceDate FROM PseudoIndexControl WHERE MailboxNumber = $MailboxNumber AND FolderId = '$($searchFolder.FolderId)'"
            $indexControlEntries = @(Get-StoreQuery -Server $db.MountedOnServer -ProcessId $db.WorkerProcessId -Query $query -Unlimited)

            # Identify the most recently accessed derived view.
            if ($indexControlEntries.Count -gt 0)
            {
                [datetime]$mostRecentViewAccessTime = [DateTime]::MinValue
                foreach ($indexControlEntry in $indexControlEntries)
                {
                    if ($indexControlEntry.LastReferenceDate -gt $mostRecentViewAccessTime)
                    {
                        $mostRecentViewAccessTime = $indexControlEntry.LastReferenceDate
                    }
                }
            }
            else
            {
                [datetime]$mostRecentViewAccessTime = $searchFolder.CreationTime
            }

            # Based on the AllowAgeOut and SearchFolderAgeOutTimeout
            # properties as well as the InstantSearch flag on the
            # search state, compute the age-out timeout (if any)
            # for the search folder and when the search folder
            # will be eligible for age-out.
            if ($ageoutTimeout -gt 0)
            {
                $friendlyAgeoutTimeout = [String]::Format("{0} ({1} seconds)", (New-TimeSpan -Seconds $ageoutTimeout), $ageoutTimeout)
                $whenEligibleForAgeout = $mostRecentViewAccessTime.AddSeconds($ageoutTimeout)
            }
            elseif ($isInstantSearch)
            {
                $friendlyAgeoutTimeout = "InstantSearch"
                $whenEligibleForAgeout = "Now"
            }
            elseif ($allowAgeout -or $searchFolder.ParentFolderId -eq $finderFolderId)
            {
                # AllowAgeout property is set, but SearchFolderAgeOutTimeout property is not set,
                # so use the default age-out timeout of 45 days (SearchFolder.TimeoutForAllowAgeOut).
                # Also, search folders under the Finder special folder are treated as if AllowAgeOut=True.
                $friendlyAgeoutTimeout = "Default (45 days)"
                $whenEligibleForAgeout = $mostRecentViewAccessTime.AddDays(45)
            }
            else
            {
                $friendlyAgeoutTimeout = "None"
                $whenEligibleForAgeout = "Never"
            }

            [PSObject]$outputSearchFolder = new-Object PSObject
            Add-Member -InputObject $outputSearchFolder -MemberType NoteProperty -Name "DisplayName" -Value $searchFolder.DisplayName
            Add-Member -InputObject $outputSearchFolder -MemberType NoteProperty -Name "FolderId" -Value $searchFolder.FolderId
            Add-Member -InputObject $outputSearchFolder -MemberType NoteProperty -Name "ParentFolderId" -Value $searchFolder.ParentFolderId
            Add-Member -InputObject $outputSearchFolder -MemberType NoteProperty -Name "CreationTime" -Value $searchFolder.CreationTime
            Add-Member -InputObject $outputSearchFolder -MemberType NoteProperty -Name "LogicalIndexNumber" -Value $searchFolder.LogicalIndexNumber
            Add-Member -InputObject $outputSearchFolder -MemberType NoteProperty -Name "MessageCount" -Value $searchFolder.MessageCount
            Add-Member -InputObject $outputSearchFolder -MemberType NoteProperty -Name "HiddenItemCount" -Value $searchFolder.HiddenItemCount
            Add-Member -InputObject $outputSearchFolder -MemberType NoteProperty -Name "FolderCount" -Value $searchFolder.FolderCount
            Add-Member -InputObject $outputSearchFolder -MemberType NoteProperty -Name "SetSearchCriteriaFlags" -Value $friendlySetSearchCriteriaFlags
            Add-Member -InputObject $outputSearchFolder -MemberType NoteProperty -Name "SearchState" -Value $friendlySearchState
            Add-Member -InputObject $outputSearchFolder -MemberType NoteProperty -Name "QueryCriteria" -Value $friendlyRestriction
            Add-Member -InputObject $outputSearchFolder -MemberType NoteProperty -Name "ScopeFolderIdList" -Value $friendlyScopeFolders
            Add-Member -InputObject $outputSearchFolder -MemberType NoteProperty -Name "IsMaterializedRestriction" -Value $($searchFolder.ParentFolderId -eq $restrictionsFolderId)
            Add-Member -InputObject $outputSearchFolder -MemberType NoteProperty -Name "IsFinder" -Value $($searchFolder.ParentFolderId -eq $finderFolderId)
            Add-Member -InputObject $outputSearchFolder -MemberType NoteProperty -Name "FolderViews" -Value $indexControlEntries.Count
            Add-Member -InputObject $outputSearchFolder -MemberType NoteProperty -Name "MostRecentViewAccessTime" -Value $mostRecentViewAccessTime
            Add-Member -InputObject $outputSearchFolder -MemberType NoteProperty -Name "AgeOutTimeout" -Value $friendlyAgeoutTimeout
            Add-Member -InputObject $outputSearchFolder -MemberType NoteProperty -Name "WhenEligibleForAgeOut" -Value $whenEligibleForAgeout

            Write-Output $outputSearchFolder

            if (++$numSearchFoldersProcessed % 50 -eq 0)
            {
                Write-Verbose "$numSearchFoldersProcessed search folders processed."
            }
        }

        Write-Verbose "Finished processing $numSearchFoldersProcessed search folders."
    }
}

# Internal helper function to determine if a property value has been redacted.
function __IsPropertyValueRedacted()
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $propertyValue
    )        

    if ([String]::IsNullOrEmpty($propertyValue))
    {
        # If string is null/empty, assume the property value wasn't redacted.
        [bool]$isRedacted = $false
    }
    else
    {
        # Check if the property value is of the form "REDACTED (x bytes)",
        # and if so, assume it was redacted.
        [int]$propertyValueLength = $propertyValue.Length
        [bool]$isRedacted = $propertyValueLength -gt 17 `
            -and $propertyValueLength -lt 27 `
            -and $propertyValue.Substring(0, 10) -eq 'REDACTED (' `
            -and $propertyValue.Substring($propertyValueLength - 7) -eq ' bytes)'
    }

    return $isRedacted
}

# Enumerates the folders in a specified mailbox and dumps the count of
# recursive and non-recursive search backlinks for each.
function Get-StoreQuerySearchBacklinkCounts
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $Database,

        [Parameter(Mandatory = $true)]
        [int]
        $MailboxNumber
    )

    $dbs = @(Get-MailboxDatabase -Identity $Database)
    if ($dbs.Count -eq 0)
    {
        Write-Error "Database $Database not found."
    }
    elseif ($dbs.Count -gt 1)
    {
        Write-Error "Database $Database matches more than one database."
    }
    else
    {
        [string]$db = $dbs[0].Name

        # Retrieve recursive and non-recursive search backlinks for all
        # folders in the specified mailbox.
        [string]$propertiesToFetch =
            "DisplayName," + `
            "FolderId," + `
            "ParentFolderId," + `
            "NonRecursiveSearchBacklinks," + `
            "RecursiveSearchBacklinks"
        [string]$query = "SELECT " + $propertiesToFetch + " FROM Folder WHERE MailboxNumber = $MailboxNumber"
        $folders = @(Get-StoreQuery -Database $db -Query $query -Unlimited)

        # Parse the recursive and non-recursive search backlinks for each
        # folder to extract the backlink counts.
        foreach ($folder in $folders)
        {
            [int]$numNonRecursiveSearchBacklinks = __ComputeNumSearchBacklinks($folder.NonRecursiveSearchBacklinks)
            [int]$numRecursiveSearchBacklinks = __ComputeNumSearchBacklinks($folder.RecursiveSearchBacklinks)

            # Compute total search backlinks. Note that we can't simply sum
            # the counts of the recursive and non-recursive search backlinks
            # because we set the count to -1 if there was some error
            # encountered attempting to parse the count out of the serialized
            # backlinks list.
            [int]$totalSearchBacklinks = 0
            if ($numNonRecursiveSearchBacklinks -gt 0)
            {
                $totalSearchBacklinks += $numNonRecursiveSearchBacklinks
            }
            if ($numRecursiveSearchBacklinks -gt 0)
            {
                $totalSearchBacklinks += $numRecursiveSearchBacklinks
            }
            
            Add-Member -InputObject $folder -MemberType NoteProperty -Name "NonRecursiveSearchBacklinksCount" -Value $numNonRecursiveSearchBacklinks
            Add-Member -InputObject $folder -MemberType NoteProperty -Name "RecursiveSearchBacklinksCount" -Value $numRecursiveSearchBacklinks
            Add-Member -InputObject $folder -MemberType NoteProperty -Name "TotalSearchBacklinksCount" -Value $totalSearchBacklinks

            Write-Output $folder
        }
    }   
}

# Internal helper function to determine the number of search backlinks given
# a serialized backlinks list.
function __ComputeNumSearchBacklinks()
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $searchBacklinks
    )        

    if ([String]::IsNullOrEmpty($searchBacklinks) -or $searchBacklinks -eq 'NULL')
    {
        [int]$numSearchBacklinks = 0
    }
    else
    {
        # First two characters are always "0x", then the next 8 characters are
        # the count of backlinks (a 4-byte integer in little-endian), and then
        # each set of 16 characters thereafter is one folder id (4 characters
        # for the 2-byte replid in little-endian form and 12 characters for
        # the 6-byte globcnt in big-endian form). So the string must be at
        # least 26 characters in length (i.e. 1 entry) and the total string
        # length minus the first 10 characters must be divisible by 16.
        [int]$searchBacklinksLength = $searchBacklinks.Length
        if ($searchBacklinksLength -lt 26 -or ($searchBacklinksLength - 10) % 16 -ne 0)
        {
            # Use -1 to indicate error.
            [int]$numSearchBacklinks = -1;
        }
        else
        {
            # CONSIDER: Should we also parse the count out of the serialized
            # string to verify it matches this computation?
            [int]$numSearchBacklinks = ($searchBacklinksLength - 10) / 16
        }
    }

    return $numSearchBacklinks
}

New-Alias -Name "fx" -Value "Format-Text" -ErrorAction SilentlyContinue

Set-StrictMode -Off


# SIG # Begin signature block
# MIIdxwYJKoZIhvcNAQcCoIIduDCCHbQCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUNggC1yxbag+ef+k2DQTlW3SA
# IR+gghhkMIIEwzCCA6ugAwIBAgITMwAAAJzu/hRVqV01UAAAAAAAnDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTMw
# WhcNMTcwNjMwMTkyMTMwWjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjU4NDctRjc2MS00RjcwMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAzCWGyX6IegSP
# ++SVT16lMsBpvrGtTUZZ0+2uLReVdcIwd3bT3UQH3dR9/wYxrSxJ/vzq0xTU3jz4
# zbfSbJKIPYuHCpM4f5a2tzu/nnkDrh+0eAHdNzsu7K96u4mJZTuIYjXlUTt3rilc
# LCYVmzgr0xu9s8G0Eq67vqDyuXuMbanyjuUSP9/bOHNm3FVbRdOcsKDbLfjOJxyf
# iJ67vyfbEc96bBVulRm/6FNvX57B6PN4wzCJRE0zihAsp0dEOoNxxpZ05T6JBuGB
# SyGFbN2aXCetF9s+9LR7OKPXMATgae+My0bFEsDy3sJ8z8nUVbuS2805OEV2+plV
# EVhsxCyJiQIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFD1fOIkoA1OIvleYxmn+9gVc
# lksuMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAFb2avJYCtNDBNG3nxss1ZqZEsphEErtXj+MVS/RHeO3TbsT
# CBRhr8sRayldNpxO7Dp95B/86/rwFG6S0ODh4svuwwEWX6hK4rvitPj6tUYO3dkv
# iWKRofIuh+JsWeXEIdr3z3cG/AhCurw47JP6PaXl/u16xqLa+uFLuSs7ct7sf4Og
# kz5u9lz3/0r5bJUWkepj3Beo0tMFfSuqXX2RZ3PDdY0fOS6LzqDybDVPh7PTtOwk
# QeorOkQC//yPm8gmyv6H4enX1R1RwM+0TGJdckqghwsUtjFMtnZrEvDG4VLA6rDO
# lI08byxadhQa6k9MFsTfubxQ4cLbGbuIWH5d6O4wggYHMIID76ADAgECAgphFmg0
# AAAAAAAcMA0GCSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAX
# BgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0MDMx
# MzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAf
# BgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJ+hbLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn
# 0UytdDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4EmPCJzB/LMySHnfL0
# Zxws/HvniB3q506jocEjU8qN+kXPCdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4n
# rIZPVVIM5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0RZCfSABKR2YR
# JylmqJfk0waBSqL5hKcRRxQJgp+E7VV4/gGaHVAIhQAQMEbtt94jRrvELVSfrx54
# QTF3zJvfO4OToWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAaswggGnMA8G
# A1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMPMAsG
# A1UdDwQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQOrIJg
# QFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcG
# CgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3Qg
# Q2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxzWPQHEy5lMFAGA1UdHwRJ
# MEcwRaBDoEGGP2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1
# Y3RzL21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYB
# BQUHMAKGOGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0Um9vdENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEB
# BQUAA4ICAQAQl4rDXANENt3ptK132855UU0BsS50cVttDBOrzr57j7gu1BKijG1i
# uFcCy04gE1CZ3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji8FMV3U+r
# kuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZLg33B+JwvBhOnY5rCnKVuKE5nGct
# xVEO6mJcPxaYiyA/4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tOi3/F
# NSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbMUVbo
# nXCUbKw5TNT2eb+qGHpiKe+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0
# NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh0EPp
# K+m79EjMLNTYMoBMJipIJF9a6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2J
# oXZhtG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWwymO0
# eFQF1EEuUKyUsKV4q7OglnUa2ZKHE3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TCCBhAwggP4
# oAMCAQICEzMAAABkR4SUhttBGTgAAAAAAGQwDQYJKoZIhvcNAQELBQAwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTAeFw0xNTEwMjgyMDMxNDZaFw0xNzAx
# MjgyMDMxNDZaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29ycG9yYXRpb24w
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCTLtrY5j6Y2RsPZF9NqFhN
# FDv3eoT8PBExOu+JwkotQaVIXd0Snu+rZig01X0qVXtMTYrywPGy01IVi7azCLiL
# UAvdf/tqCaDcZwTE8d+8dRggQL54LJlW3e71Lt0+QvlaHzCuARSKsIK1UaDibWX+
# 9xgKjTBtTTqnxfM2Le5fLKCSALEcTOLL9/8kJX/Xj8Ddl27Oshe2xxxEpyTKfoHm
# 5jG5FtldPtFo7r7NSNCGLK7cDiHBwIrD7huTWRP2xjuAchiIU/urvzA+oHe9Uoi/
# etjosJOtoRuM1H6mEFAQvuHIHGT6hy77xEdmFsCEezavX7qFRGwCDy3gsA4boj4l
# AgMBAAGjggF/MIIBezAfBgNVHSUEGDAWBggrBgEFBQcDAwYKKwYBBAGCN0wIATAd
# BgNVHQ4EFgQUWFZxBPC9uzP1g2jM54BG91ev0iIwUQYDVR0RBEowSKRGMEQxDTAL
# BgNVBAsTBE1PUFIxMzAxBgNVBAUTKjMxNjQyKzQ5ZThjM2YzLTIzNTktNDdmNi1h
# M2JlLTZjOGM0NzUxYzRiNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzcitW2oynUC
# lTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# b3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEGCCsGAQUF
# BwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0MAwGA1Ud
# EwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAIjiDGRDHd1crow7hSS1nUDWvWas
# W1c12fToOsBFmRBN27SQ5Mt2UYEJ8LOTTfT1EuS9SCcUqm8t12uD1ManefzTJRtG
# ynYCiDKuUFT6A/mCAcWLs2MYSmPlsf4UOwzD0/KAuDwl6WCy8FW53DVKBS3rbmdj
# vDW+vCT5wN3nxO8DIlAUBbXMn7TJKAH2W7a/CDQ0p607Ivt3F7cqhEtrO1Rypehh
# bkKQj4y/ebwc56qWHJ8VNjE8HlhfJAk8pAliHzML1v3QlctPutozuZD3jKAO4WaV
# qJn5BJRHddW6l0SeCuZmBQHmNfXcz4+XZW/s88VTfGWjdSGPXC26k0LzV6mjEaEn
# S1G4t0RqMP90JnTEieJ6xFcIpILgcIvcEydLBVe0iiP9AXKYVjAPn6wBm69FKCQr
# IPWsMDsw9wQjaL8GHk4wCj0CmnixHQanTj2hKRc2G9GL9q7tAbo0kFNIFs0EYkbx
# Cn7lBOEqhBSTyaPS6CvjJZGwD0lNuapXDu72y4Hk4pgExQ3iEv/Ij5oVWwT8okie
# +fFLNcnVgeRrjkANgwoAyX58t0iqbefHqsg3RGSgMBu9MABcZ6FQKwih3Tj0DVPc
# gnJQle3c6xN3dZpuEgFcgJh/EyDXSdppZzJR4+Bbf5XA/Rcsq7g7X7xl4bJoNKLf
# cafOabJhpxfcFOowMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkqhkiG9w0B
# AQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEyMDAG
# A1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IDIwMTEw
# HhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBT
# aWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA
# q/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03a8YS2Avw
# OMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akrrnoJr9eW
# WcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0RrrgOGSsbmQ1
# eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy4BI6t0le
# 2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9sbKvkjh+
# 0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAhdCVfGCi2
# zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8kA/DRelsv
# 1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTBw3J64HLn
# JN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmnEyimp31n
# gOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90lfdu+Hgg
# WCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0wggHpMBAG
# CSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2oynUClTAZ
# BgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/
# BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBaBgNVHR8E
# UzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9k
# dWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsGAQUFBwEB
# BFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9j
# ZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNVHSAEgZcw
# gZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsGAQUFBwIC
# MDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABlAG0AZQBu
# AHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKbC5YR4WOS
# mUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11lhJB9i0ZQ
# VdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6I/MTfaaQ
# dION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0wI/zRive
# /DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560STkKxgrC
# xq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQamASooPoI/
# E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGaJ+HNpZfQ
# 7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ahXJbYANah
# Rr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA9Z74v2u3
# S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33VtY5E90Z1W
# Tk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr/Xmfwb1t
# bWrJUnMTDXpQzTGCBM0wggTJAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB4TAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUO631Col/q8lVmbB7/4nStlXMpPwwgYAGCisG
# AQQBgjcCAQwxcjBwoEiARgBNAGEAbgBhAGcAZQBkAFMAdABvAHIAZQBEAGkAYQBn
# AG4AbwBzAHQAaQBjAEYAdQBuAGMAdABpAG8AbgBzAC4AcABzADGhJIAiaHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQAs
# MkB1TjskkY6r5dJhLhdoZdGFrWcT2wgdoBCG1aupWPBZOpKhorZekVB4gx8+uanj
# rf4Y38l1btvXMsNkINmCkYemwg6MEvRo7yiQq6BMOzgNzDy1iqSAacnAzsbUc9Oh
# ayT6lge0nQfyC6sI3dNMqTVAcVmyBBLmr26Vskh5dItRc/SV6I+qfcLAuaNRvpGd
# PR43zBvVi/u4AiYAzz8cozAw6CZI4krEgl0iqHVsuxvvGBak5FzGBo+1jqcKd3h+
# vPUNmAkPASoP8hI2tQdNTRannLoEJuRRuoUL8wqAbdsSmr1zjJFrpCfhHvdYooSw
# NbAmg8UAJuejYB4IjZvIoYICKDCCAiQGCSqGSIb3DQEJBjGCAhUwggIRAgEBMIGO
# MHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMT
# GE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQQITMwAAAJzu/hRVqV01UAAAAAAAnDAJ
# BgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0B
# CQUxDxcNMTYwOTAzMTg0NTAyWjAjBgkqhkiG9w0BCQQxFgQUHR6e73MZAWREAU2F
# 1RyWJ6OT5XIwDQYJKoZIhvcNAQEFBQAEggEAbj0fdj/OleSrp0wpZ1fQpgVV5Av9
# dhic7AvScozFI8IIqorjNn2CZVxZFxjYH4LFcaJ8mZ7iE11bV/+iTmJCqo8eBDzq
# 0/x45g/YohwKpYMiw7NxByu/yeXYKIGvkSvFzvZuEdoQhCGbY3H+lgyiiJ8Tca44
# OA1iUQmqGeg7sagHhNk8KVkZiE2lgygsv10eKpgGUeB/+AAg+g3oYiFHqZhlNf+0
# dj1feN/bDRu5IIXl+8Dz80uy2aOe6UxKmGbst3xx6XYw8D3Iz9HMw9GhgN0m3wf8
# oQ2xA7UpCs9PO3xbkw0twfUd9Y67ReciVppWiRvFwLdNScKqvlpqATEKtA==
# SIG # End signature block

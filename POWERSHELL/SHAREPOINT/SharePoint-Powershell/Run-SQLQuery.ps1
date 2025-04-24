# Example: 
# Run-SQLQuery -SqlServer "SQLSERVER" -SqlDatabase "SharePoint_Content_Portal" `
# -SqlQuery "SELECT * from AllDocs where SetupPath = 'Features\ReviewPageListInstances\Files\Workflows\ReviewPage\ReviewPage.xoml'" |`
# select Id, SiteId, DirName, LeafName, WebId, ListId | Format-List

function Run-SQLQuery ($SqlServer, $SqlDatabase, $SqlQuery)
{
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server =" + $SqlServer + "; Database =" + $SqlDatabase + "; Integrated Security = True"
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $SqlCmd.CommandText = $SqlQuery
    $SqlCmd.Connection = $SqlConnection
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlAdapter.SelectCommand = $SqlCmd
    $DataSet = New-Object System.Data.DataSet
    $SqlAdapter.Fill($DataSet)
    $SqlConnection.Close()
    $DataSet.Tables[0]
}
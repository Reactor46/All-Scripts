
#Author Vince Vardaro

$head = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html><head><title>Users In Active Directory Report</title>
<style type="text/css">
<!--
body {
font-family: Tahoma, Verdana, Arial;
}
  table
    {
    
              margin-right:auto;
        border: 1px solid rgb(190, 190, 190);
        font-Family: Helvetica, Tahoma;
        font-Size: 10pt;
        text-align: left;
    }
th
    {
        Text-Align: Left;
        font-size: 8pt;
        font-weight: bold;
        Color:#4682B4;
              background-color: white;
        Padding: 1px 4px 1px 4px;
    }
tr:hover td
    {
        background-color: DodgerBlue ;
        Color: #F5FFFA;
           
    }
tr:nth-child(even)
    {
        Background-Color: #D3D3D3;
    }
tr:nth-child(odd)
       {
              background-color:#F8F8FF;
       }   
    
td
    {
        Vertical-Align: Top;
        Padding: 1px 4px 1px 4px;
    }
 
h1{
       clear: both;
       font-size: 12pt;
       font-weight: bold;
       }
h2{
       clear: both;
    
    
       font-size: 17px;
       font-weight: 300;
       
}
p{
    font-size: 10pt;
       font-weight: 300; 
    text-align: left;
    margin-bottom: 10px;
}
}
-->
</style>
</head>
"@
$DateStamp = Get-Date
$Domain = Get-ADDomain -Server LASDC02 |select -ExpandProperty NetBIOSName
$Users = Get-ADUser -Filter * -Properties *
$count = $Users.count
$EnabledCount = ($Users|where {$_.Enabled -eq $true}).Count
$DisabledCount = ($Users|where {$_.Enabled -eq $false}).Count
$30days = ($Users|Where {(((Get-date) - ([datetime]::FromFileTime($_.lastlogontimestamp))).TotalDays -gt 30)}).Count
$60days = ($Users|Where {(((Get-date) - ([datetime]::FromFileTime($_.lastlogontimestamp))).TotalDays -gt 60)}).Count
$90days = ($Users|Where {(((Get-date) - ([datetime]::FromFileTime($_.lastlogontimestamp))).TotalDays -gt 90)}).Count
#$GT90days = ($Users|where { $_.LastLogonDate.AddDays($90) -lt $CurrentDate }).Count
$Users|Sort-Object -property name |
ConvertTo-Html Enabled,Name,SamAccountName,UserPrincipalName,WhenCreated,LastLogonDate,LogonCount -head $head `
-body "<body><h3> $Domain AD Users</h3>
       <h3> Report Last Run: $DateStamp</h3>
       <p>Total Users = $count</p>
       <p style=`"color: rgb(41, 138, 8);`">Active Users: $EnabledCount</p>
       <p style=`"color: rgb(164, 0, 0);`">Disabled Users: $DisabledCount</p>
       <p style=`"color: rgb(164, 0, 0);`">Accounts not accessed in 30 Days: $30Days</p>
       <p style=`"color: rgb(164, 0, 0);`">Accounts not accessed in 60 Days: $60Days</p>
       <p style=`"color: rgb(164, 0, 0);`">Accounts not accessed in 90 Days: $90Days</p></body>" |
Out-File "C:\Scripts\Repository\jbattista\Web\reports\$($Domain)UserReport.html"
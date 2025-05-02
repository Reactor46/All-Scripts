$servers = Get-Content D:\IIS-Servers\simple_inventory_IIS\Servers.txt
#psexec -accepteula @D:\IIS-Servers\simple_inventory_IIS\Servers.txt -s winrm.cmd quickconfig -q

## Creating HTML header ##
"<!DOCTYPE html>
<HTML>
<div align=""center"">
<body>
<font color =""black"" face=""Microsoft Tai le"">
<H2> IIS Inventory </H2>" > index.html

## Creating column titles in HTML ##
"<Table border=1 cellpadding=0 cellspacing=0>
<TR bgcolor=red align=center>
<TD><FONT COLOR=""white""><B>&nbsp; HOSTNAME &nbsp</B></TD>
<TD><FONT COLOR=""white""><B>&nbsp; OS &nbsp</B></TD>
<TD><FONT COLOR=""white""><B>&nbsp; IP &nbsp</B></TD>
<TD><FONT COLOR=""white""><B>&nbsp; IIS_Version &nbsp</B></TD>
<TD><FONT COLOR=""white""><B>&nbsp; SITES &nbsp</B></TD>
<TD><FONT COLOR=""white""><B>&nbsp; APPLICATIONS POOL &nbsp</B></TD>
<TD><FONT COLOR=""white""><B>&nbsp; APPLICATIONS &nbsp</B></TD>
<TD><FONT COLOR=""white""><B>&nbsp; VDIR &nbsp</B></TD>
</TR>" >> index.html

## Start of loop for list of servers ##
foreach($server in $servers){
    ## Collect variables of OS IIS version and IP##
    $vIIS = Invoke-Command -ComputerName $server -scriptblock {(Get-ItemProperty HKLM:\SOFTWARE\Microsoft\InetStp\ | Select-Object setupstring | Format-Wide | Out-String).Trim()}
    $SO = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $server | ForEach-Object -MemberName Caption
    $ip = (Test-Connection -ComputerName $server -Count 1 | Select-Object IPV4Address | Format-Wide | Out-String).Trim()

    ########################################## ApplicationPool ##########################################
    $lista_pool = invoke-command -computername $server -scriptblock {C:\Windows\System32\inetsrv\.\appcmd.exe list apppool} #| ForEach-Object{$_.Split('"')[1];}
    $count_pool = $lista_pool.Count

    $html ="<!DOCTYPE>
    <HTML><div align=""center"">
    <body>
    <font color =""black"" face=""Microsoft Tai le""><H2> Applications Pools </H2>"

    $html +="<Table border=1 cellpadding=0 cellspacing=0>
    <TR bgcolor=red align=center>
    <TD><FONT COLOR=""white""><B>&nbsp; Server &nbsp</B></TD>
    <TD><FONT COLOR=""white""><B>&nbsp; Results &nbsp</B></TD>
    </TR>"

    foreach ($pool in $lista_pool){
        $html += "<TR>
        <TD align=center>&nbsp;$server&nbsp</TD>
        <TD align=center>&nbsp;$pool&nbsp</TD>
        </TR>"
    }

    $html | Out-File .\pool\pool_$server.html

    ########################################## Site ##########################################
    $site_list = invoke-command -computername $server -scriptblock {C:\Windows\System32\inetsrv\.\appcmd.exe list site} #| ForEach-Object{$_.Split('"')[1,2];}
    #Invoke-Command -ComputerName $server -ScriptBlock {Import-Module WebAdministration}
    #$s = New-PSSession -ComputerName $server
    #Invoke-Command -Session $s -ScriptBlock { Import-Module WebAdministration }
    #$site_list = Invoke-Command -Session $s -ScriptBlock { Get-ChildItem -Path IIS:\Sites } | ForEach-Object{$_.Split('"')[1];}
    #$site_list = @(Get-WebBinding | % {
    #$name = $_.ItemXPath -replace '(?:.*?)name=''([^'']*)(?:.*)', '$1'
    #New-Object psobject -Property @{
    #    Name = $name
    #    Binding = $_.bindinginformation.Split(":")[-1]    }} | Group-Object -Property Name | 
    #Format-Table Name, @{n="Bindings";e={$_.Group.Binding -join "`n"}} -Wrap )
    #Invoke-Command -Session $s -ScriptBlock { $site_list }
    $site_count = $site_list.Count

    $html ="<!DOCTYPE>
    <HTML><div align=""center"">
    <body>
    <font color =""black"" face=""Microsoft Tai le""><H2> Sites </H2>"

    $html +="<Table border=1 cellpadding=0 cellspacing=0>
    <TR bgcolor=red align=center>
    <TD><FONT COLOR=""white""><B>&nbsp; Server &nbsp</B></TD>
    <TD><FONT COLOR=""white""><B>&nbsp; Site_Name  &nbsp</B></TD>
    </TR>"

    foreach ($site in $site_list){
        $html += "<TR>
        <TD align=center>&nbsp;$server&nbsp</TD>
        <TD align=center>&nbsp;$site&nbsp</TD>
        </TR>"
    }
    #Get-PSSession | Remove-PSSession
    $html | Out-File .\site\site_$server.html

    ########################################## Application ##########################################
    $app_list = invoke-command -computername $server -scriptblock {C:\Windows\System32\inetsrv\.\appcmd.exe list app} #| ForEach-Object{$_.Split('"')[1];}
    $app_count = $app_list.Count

    $html ="<!DOCTYPE>
    <HTML><div align=""center"">
    <body>
    <font color =""black"" face=""Microsoft Tai le""><H2> Application </H2>"
    
    $html +="<Table border=1 cellpadding=0 cellspacing=0>
    <TR bgcolor=red align=center>
    <TD><FONT COLOR=""white""><B>&nbsp; Servidor &nbsp</B></TD>
    <TD><FONT COLOR=""white""><B>&nbsp; Resultados encontrados &nbsp</B></TD>
    </TR>"

    foreach ($app in $app_list){
        $html += "<TR>
        <TD align=center>&nbsp;$server&nbsp</TD>
        <TD align=center>&nbsp;$app&nbsp</TD>
        </TR>"
    }
        
    $html | Out-File .\app\app_$server.html

    ########################################## Diretorio Virtual ##########################################
    $vdir_list = invoke-command -computername $server -scriptblock {C:\Windows\System32\inetsrv\.\appcmd.exe list vdir /config} #| ForEach-Object{$_.Split('"')[3];}
    $vdir_count = $vdir_list.Count

    $html ="<!DOCTYPE>
    <HTML><div align=""center"">
    <body>
    <font color =""black"" face=""Microsoft Tai le""><H2> Diretorios Virtuais </H2>"

    $html +="<Table border=1 cellpadding=0 cellspacing=0>
    <TR bgcolor=red align=center>
    <TD><FONT COLOR=""white""><B>&nbsp; Servidor &nbsp</B></TD>
    <TD><FONT COLOR=""white""><B>&nbsp; Resultados encontrados &nbsp</B></TD>
    </TR>"

    foreach ($vdir in $vdir_list){
        $html += "<TR>
        <TD align=center>&nbsp;$server&nbsp</TD>
        <TD align=center>&nbsp;$vdir&nbsp</TD>
        </TR>"
    }

    $html | Out-File .\vdir\vdir_$server.html

    ########################################## INDEX ##########################################
    "<TD align=center>&nbsp; $server &nbsp</TD>
    <TD align=center>&nbsp; $SO &nbsp</TD>
    <TD align=center>&nbsp; $ip &nbsp</TD>
    <TD align=center>&nbsp; $vIIS &nbsp</TD>
    <TD align=center><a href="".\site\site_$server.html"">$site_count</TD>
    <TD align=center><a href="".\pool\pool_$server.html"">$pool_count</TD>
    <TD align=center><a href="".\app\app_$server.html"">$app_count</TD>
    <TD align=center><a href="".\vdir\vdir_$server.html"">$vdir_count</TD>
    </TR>" >> index.html
}
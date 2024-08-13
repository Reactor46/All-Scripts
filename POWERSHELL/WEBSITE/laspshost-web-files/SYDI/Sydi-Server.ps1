Import-Module C:\Scripts\Repository\jbattista\SYDI\Out-FileUtf8NoBom.ps1

$date = Get-Date -Format "M-dd-yyyy"
#$date = "3-08-2018"
$baseDIR = "C:\Scripts\Repository\jbattista\SYDI"
$output_Dir = "\\lasfs02\Winsys$\SYDI_Server_Info\Output_files"

# Create Dirs
New-Item -Path "$output_Dir\$date\TXT-Results" -ItemType Directory -Force
New-Item -Path "$output_Dir\$date\Contoso.CORP" -ItemType Directory -Force
New-Item -Path "$output_Dir\$date\Contoso.CORP\DOC" -ItemType Directory -Force
New-Item -Path "$output_Dir\$date\Contoso.CORP\XML" -ItemType Directory -Force
New-Item -Path "$output_Dir\$date\Contoso.CORP\HTML" -ItemType Directory -Force
New-Item -Path "$output_Dir\$date\PHX.Contoso.CORP" -ItemType Directory -Force
New-Item -Path "$output_Dir\$date\PHX.Contoso.CORP\DOC" -ItemType Directory -Force
New-Item -Path "$output_Dir\$date\PHX.Contoso.CORP\XML" -ItemType Directory -Force
New-Item -Path "$output_Dir\$date\PHX.Contoso.CORP\HTML" -ItemType Directory -Force
New-Item -Path "$output_Dir\$date\CREDITONEAPP.TST" -ItemType Directory -Force
New-Item -Path "$output_Dir\$date\CREDITONEAPP.TST\DOC" -ItemType Directory -Force
New-Item -Path "$output_Dir\$date\CREDITONEAPP.TST\XML" -ItemType Directory -Force
New-Item -Path "$output_Dir\$date\CREDITONEAPP.TST\HTML" -ItemType Directory -Force
New-Item -Path "$output_Dir\$date\CREDITONEAPP.BIZ" -ItemType Directory -Force
New-Item -Path "$output_Dir\$date\CREDITONEAPP.BIZ\DOC" -ItemType Directory -Force
New-Item -Path "$output_Dir\$date\CREDITONEAPP.BIZ\XML" -ItemType Directory -Force
New-Item -Path "$output_Dir\$date\CREDITONEAPP.BIZ\HTML" -ItemType Directory -Force


# Find Servers
Get-ADComputer -Server LASDC02.Contoso.corp -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom $output_Dir\$date\TXT-Results\Contoso.TXT  -Append 
Get-ADComputer -Server PHXDC03.phx.Contoso.corp -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom $output_Dir\$date\TXT-Results\PHX.TXT  -Append
Get-ADComputer -Server LASAUTHTST01.creditoneapp.tst -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom $output_Dir\$date\TXT-Results\C1B-TST.TXT  -Append 
Get-ADComputer -Server LASAUTH01.creditoneapp.biz -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty DNSHostName | Out-FileUtf8NoBom $output_Dir\$date\TXT-Results\C1B-BIZ.TXT  -Append 


# Check alive or dead
Get-Content $output_Dir\$date\TXT-Results\Contoso.TXT | 

 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom $output_Dir\$date\TXT-Results\Contoso.Alive.txt -append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom $output_Dir\$date\TXT-Results\Contoso.Dead.txt -append}}

  Get-Content $output_Dir\$date\TXT-Results\PHX.TXT | 

 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom $output_Dir\$date\TXT-Results\PHX.Alive.txt -append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom $output_Dir\$date\TXT-Results\PHX.Dead.txt -append}}

  Get-Content $output_Dir\$date\TXT-Results\C1B-TST.TXT | 

 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom $output_Dir\$date\TXT-Results\C1B-TST.Alive.txt -append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom $output_Dir\$date\TXT-Results\C1B-TST.Dead.txt -append}}

  Get-Content $output_Dir\$date\TXT-Results\C1B-BIZ.TXT | 

 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom $output_Dir\$date\TXT-Results\C1B-BIZ.Alive.txt -append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom $output_Dir\$date\TXT-Results\C1B-BIZ.Dead.txt -append}}


# Write-XML
$Servers = Get-Content $output_Dir\$date\TXT-Results\Contoso.Alive.txt
    ForEach($srv in $Servers){
    cscript /nologo $baseDIR\sydi-server.vbs -ex -t"$srv" -o"$output_Dir\$date\Contoso.CORP\XML\$srv.xml" -sh
    cscript /nologo $baseDIR\tools\ss-xml2word.vbs -d -x"$output_Dir\$date\Contoso.CORP\XML\$srv.xml" -l"$baseDIR\tools\lang_english.xml" -o"$output_Dir\$date\Contoso.CORP\DOC\$srv.doc"
    cscript /nologo $baseDIR\tools\sydi-transform.vbs" -x$output_Dir\$date\Contoso.CORP\XML\$srv.xml" -s"$baseDIR\xml\serverhtml.xsl" -o"$output_Dir\$date\Contoso.CORP\HTML\$srv.html"

    }

$Servers2 = Get-Content $output_Dir\$date\TXT-Results\PHX.Alive.txt
    ForEach($srv2 in $Servers2){
    cscript /nologo $baseDIR\sydi-server.vbs -ex -t"$srv2" -o"$output_Dir\$date\PHX.Contoso.CORP\XML\$srv2.xml" -sh

    }
    
$Servers3 = Get-Content $output_Dir\$date\TXT-Results\C1B-BIZ.Alive.txt
    ForEach($srv3 in $Servers3){
    cscript /nologo $baseDIR\sydi-server.vbs -ex -t"$srv3" -o"$output_Dir\$date\CREDITONEAPP.BIZ\XML\$srv3.xml" -sh

    }

$Servers4 = Get-Content $output_Dir\$date\TXT-Results\C1B-TST.Alive.txt
    ForEach($srv4 in $Servers4){
    cscript /nologo $baseDIR\sydi-server.vbs -ex -t"$srv4" -o"$output_Dir\$date\CREDITONEAPP.TST\XML\$srv4.xml" -sh

    }
  
# Convert from XML to DOC and HTML
$ContosoXML = "$output_Dir\$date\Contoso.CORP\XML"
$PHXXML = "$output_Dir\$date\PHX.Contoso.CORP\XML"
$TSTXML = "$output_Dir\$date\CREDITONEAPP.TST\XML"
$BIZXML = "$output_Dir\$date\CREDITONEAPP.BIZ\XML"

gci $ContosoXML\|sort |`
    foreach {
$file1 = $_.name
$fileBasename1 = $_.basename

cscript /nologo $baseDIR\tools\ss-xml2word.vbs -d -x"$ContosoXML\$file1" -l"$baseDIR\tools\lang_english.xml" -o"$output_Dir\$date\Contoso.CORP\DOC\$fileBasename1.doc"
cscript /nologo $baseDIR\tools\sydi-transform.vbs" -x"$ContosoXML\$file1" -s"$baseDIR\xml\serverhtml.xsl" -o"$output_Dir\$date\Contoso.CORP\HTML\$fileBasename1.html"
}

gci $PHXXML\|sort |`
    foreach {
$file2 = $_.name
$fileBasename2 = $_.basename

cscript "$baseDIR\tools\ss-xml2word.vbs" "-d" "-x$PHXXML\$file2" "-l$baseDIR\tools\lang_english.xml" "-s$output_Dir\$date\PHX.Contoso.CORP\DOC\$fileBasename2.doc"
cscript “$baseDIR\tools\sydi-transform.vbs” "-x$PHXXML\$file2” "-s$baseDIR\xml\serverhtml.xsl" "-o$output_Dir\$date\PHX.Contoso.CORP\HTML\$fileBasename2.html"
}

gci $TSTXML\|sort |`
    foreach {
$file3 = $_.name
$fileBasename3 = $_.basename

cscript "$baseDIR\tools\ss-xml2word.vbs" "-d" "-x$file3" "-l$baseDIR\tools\lang_english.xml" "-s$output_Dir\$date\CREDITONEAPP.TST\DOC\$fileBasename3.doc"
cscript “$baseDIR\tools\sydi-transform.vbs” "-x$file3” "-s$output_Dir\xml\serverhtml.xsl" "-o$output_Dir\$date\CREDITONEAPP.TST\HTML\$fileBasename3.html"

}

gci $BIZXML\|sort |`
    foreach {
$file4 = $_.name
$fileBasename4 = $_.basename

cscript "$baseDIR\tools\ss-xml2word.vbs" "-d" "-x$file4" "-l$baseDIR\tools\lang_english.xml" "-s$output_Dir\$date\CREDITONEAPP.BIZ\DOC\$fileBasename4.doc"
cscript “$baseDIR\tools\sydi-transform.vbs” "-x$file4” "-s$baseDIR\xml\serverhtml.xsl" "-o$output_Dir\$date\CREDITONEAPP.BIZ\HTML\$fileBasename4.html"

}
#Set Variables
$vmhosts = import-csv import-csv C:\LazyWinAdmin\VMWARE\Certs\HOST-SSL.csv
$ca = vvrcrtwa01.res.vegas.com\VegasEntCA
$template = "CertificateTemplate:Web"
 
foreach($vmhost in $vmhosts){
    $name = $vmhost.Hostname
    $dirpath = "c:\certs\hosts\" + $name
    new-item -ItemType directory -path $dirpath
    $path = "c:\certs\hosts\" + $name + "\"
    $csr = "c:\openssl\bin\openssl.exe req -new -nodes -out " + $path + "rui.csr -keyout " + $path + "rui-orig.key -config c:\certs\hosts\opensslcfg\" + $name + ".cfg"
    $key = "c:\openssl\bin\openssl.exe rsa -in " + $path + "rui-orig.key " + "-out " + $path + "rui.key"
    $reqcert = "C:\windows\system32\certreq.exe -config " + $ca + " -attrib " + $template + " " + $path + "rui.csr " + $path + "rui.crt"
    IEX $csr | out-null
    IEX $key | out-null
    IEX $reqcert | out-null
    }
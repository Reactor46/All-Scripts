#Set Variables
$vmhosts = import-csv "C:\LazyWinAdmin\VMWARE\Certs\HOST-SSL.csv"
$ca = """vvrcrtwa01.res.vegas.com\VegasEntCA"""
$template = "CertificateTemplate:vSphere 6.x"
 
foreach($vmhost in $vmhosts){
    $name = $vmhost.HostName
    $dirpath = "C:\LazyWinAdmin\VMWARE\Certs\Hosts\" + $name
    new-item -ItemType directory -path $dirpath
    $path = "C:\LazyWinAdmin\VMWARE\Certs\Hosts\" + $name + "\"
    $csr = "C:\Program Files\OpenSSL-Win64\bin\openssl.exe req -new -nodes -out " + $path + "rui.csr -keyout " + $path + "rui-orig.key -config C:\LazyWinAdmin\VMWARE\Certs\Hosts\OpenSSLCfg\" + $name + ".cfg"
    $key = "C:\Program Files\OpenSSL-Win64\bin\openssl.exe rsa -in " + $path + "rui-orig.key " + "-out " + $path + "rui.key"
    $reqcert = "C:\windows\system32\certreq.exe -config " + $ca + " -attrib " + $template + " " + $path + "rui.csr " + $path + "rui.crt"
    IEX $csr | out-null
    IEX $key | out-null
    IEX $reqcert | out-null
    }
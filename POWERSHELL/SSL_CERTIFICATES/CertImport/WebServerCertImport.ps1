Read-Host -assecurestring | convertfrom-securestring | out-file .\pwd.txt -Force

. .\Import-STPfxCertificate.ps1

$Servers = GC .\TEST.txt
$CertPath = '.\star.creditone.com.pfx'
$Creds = GC .\pwd.txt | ConvertTo-SecureString
$CertRoot = 'localmachine'
$CertStore = 'My'
$Flags = 'PersistKeySet,MachineKeySet'

ForEach($Srv in $Servers){
Import-STPfxCertificate -ComputerName $Srv -Password $Creds -CertFilePath $CertPath -CertRootStore $CertRoot -CertStore $CertStore -X509Flags $Flags -ErrorAction Continue
}
    
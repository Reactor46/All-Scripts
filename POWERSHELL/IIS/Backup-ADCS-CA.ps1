Function Backup-CAConfig {
param (
[Parameter(Mandatory=$true)]
[string]$BackupPath
)
New-Item -Path $BackupPath -ItemType directory
certutil –backupdb $BackupPath
certutil -backupkey $BackupPath
certutil –getreg ca > $BackupPath\CA_certutil_getreg.txt
reg export HKLM\SYSTEM\CurrentControlSet\Services\CertSvc\Configuration $BackupPath\CA_regedir_CertSvcConfiguration.reg
Get-CATemplate|foreach{$_.Name}|out-file -filepath $BackupPath\CATemplates.txt –encoding string –force
certutil –v -catemplates > $BackupPath\Certutil_CATemplates.txt
certutil -cainfo > $BackupPath\Certutil_cainfo.txt
certutil –getreg ca\csp > $BackupPath\certutil_cacsp_getreg.txt
}
$path = < локальный путь к сертификату / local path to certificate >;
$pass = < пароль к сертификату / cetificate password >;
$subject = < subject сертификата / subject certificate >;
$date = Get-Date;
$list = Get-ExchangeCertificate;

Import-ExchangeCertificate -FileName $path -Password (ConvertTo-SecureString -string $pass -AsPlainText -Force);

ForEach($cert in $list)
{
  if ( $cert.Subject -eq $subject -and $cert.NotAfter -gt $date)
  {
    write($cert.Subject);
    write($cert.Thumbprint);
    $thumbprint = $cert.Thumbprint;
    Enable-ExchangeCertificate -Thumbprint $thumbprint -Services POP,IMAP,IIS,SMTP
    break;
  }
}
$Telegraf = 'C:\telegraf\telegraf.exe'
$Directory = $Telegraf | split-path -Parent
Start-Process -FilePath $Telegraf -ArgumentList 'config > new_telegaf_config.conf'   -WorkingDirectory $Directory  -RedirectStandardOutput ([io.path]::Combine($Directory,'telegraf.conf'))
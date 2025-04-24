$ServiceAccount = (read-host 'Enter Service Account')


#this assumes your config/install are in c:\telegraf
ii 'C:\telegraf\telegraf.conf'

#setup statements to automate install 
$EXE = 'C:\telegraf\telegraf.exe'
$Config = 'C:\telegraf\telegraf.conf'

#where to copy config to
$Destination = "C:\Program Files\telegraf"
$NewExePath = [io.path]::Combine($Destination,'telegraf.exe')
New-item -Path $Destination -ItemType Directory -Force
$Config | Copy-Item -Destination $Destination -Force -Verbose
$EXE | Copy-Item -Destination $Destination -Force -Verbose


#if already installed
#stop-Service -Name telegraf
#Start-Process -FilePath $NewExePath -ArgumentList '--service uninstall'

Start-Process -FilePath $NewExePath -ArgumentList '--service install'

#change credentials
$service = gwmi win32_service -filter "name='telegraf'"
$service.change($null,$null,$null,$null,$null,$null,$ServiceAccount,(Read-host 'Enter Password'))


get-service -name telegraf | Format-Table -AutoSize
start-service -name telegraf 

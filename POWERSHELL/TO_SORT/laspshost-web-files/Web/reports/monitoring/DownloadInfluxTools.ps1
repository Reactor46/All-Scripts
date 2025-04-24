
clear-host

[datetime]$StepTimer = [datetime]::Now
$url = 'https://portal.influxdata.com/downloads'
$DownloadDirectory = 'C:\Influx'

if(!([io.directory]::Exists($DownloadDirectory)))
{New-Item -Path 'C:\Influx' -ItemType directory -Force -ErrorAction Ignore -Verbose}

$PageHtml = Invoke-WebRequest -Uri $url 

$PageHtml | select-string -Pattern '(https\:\/\/.*?\.\d{1,2}_windows_amd64.zip)|(https:\/\/dl\.influxdata\.com\/chronograf.*?_windows_amd64.zip)' -AllMatches | select -expand Matches|  % { 
        $NewFileName = $_.Groups[0] | select-string -Pattern '(?<=\/)\w*?(?=\-)'-AllMatches |  select -ExpandProperty Matches -first 1 
        $DownloadUrl = $_.Groups[0].Value
        $DownloadedFile = ([io.path]::Combine($DownloadDirectory,"$NewFileName.zip"))
        $UnzipLocation = ([io.path]::Combine($DownloadDirectory,$NewFileName)) 
        write-debug "--- Downloading ---`nParsed Name: $($NewFIleName)`nDownload URL: $($DownloadURL)`nUnzipLocation: $($UnzipLocation)"

        Invoke-WebRequest -Uri $DownloadUrl -OutFile ([io.path]::Combine($DownloadDirectory,"$NewFileName.zip")) -Verbose
       
        if([io.directory]::Exists($DownloadDirectory))
        {Remove-Item -Path $UnzipLocation -Force -ErrorAction Ignore -Verbose -recurse}

        Expand-Archive -Path $DownloadedFile -DestinationPath ([io.path]::Combine($DownloadDirectory))  -Force
        Remove-Item -Path $DownloadedFile -Force -Verbose
      
}

Invoke-WebRequest -Uri 'https://github.com/influxdata/telegraf/archive/master.zip' -OutFile ([io.path]::Combine($DownloadDirectory,"Telegraf.zip")) -Verbose
Expand-Archive -Path ([io.path]::Combine($DownloadDirectory,"Telegraf.zip")) -DestinationPath ([io.path]::Combine($DownloadDirectory))  -Force
Remove-Item -Path ([io.path]::Combine($DownloadDirectory,"Telegraf.zip")) -Force -Verbose
      

write-debug( "{0:hh\:mm\:ss\.fff} {1}: finished" -f [timespan]::FromMilliseconds(((Get-Date)-$StepTimer).TotalMilliseconds),'Finished download and unzip')
ii $DownloadDirectory
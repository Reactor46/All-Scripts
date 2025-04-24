# SharePoint Warmup Script
# by Ingo Karstein
# 2011/01/26
# 2011/08/02

#<---- Improvement starts here
#region MyWebClient
    Add-Type -ReferencedAssemblies "System.Net" -TypeDefinition @"
    using System.Net;

    public class MyWebClient : WebClient
    {
        private int timeout = 60000;
            
        public MyWebClient(int timeout)
        {
            this.timeout = timeout;
        }
        
        public int Timeout
        {
            get
            {
                return timeout;
            }
            set
            {
                timeout = value;
            }
        }
        
        protected override WebRequest GetWebRequest(System.Uri webUrl)
        {
            WebRequest retVal = base.GetWebRequest(webUrl);
            retVal.Timeout = this.timeout;
            return retVal;
        }
    }
"@
#endregion
#----> Improvement ends here

$urls= @("http://sharepoint.local", "http://another.sharepoint.local")

New-EventLog -LogName "Application" -Source "SharePoint Warmup Script" -ErrorAction SilentlyContinue | Out-Null

$timeout = 60000 #=60 seconds               #<--- Improvement!

$urls | % {
    $url = $_
    try {
        $wc = New-Object MyWebClient($timeout)        #<--- Improvement!
        $wc.Credentials = [System.Net.CredentialCache]::DefaultCredentials
        $ret = $wc.DownloadString($url)
        if( $ret.Length -gt 0 ) {
            $s = "Last run successful for url ""$($url)"": $([DateTime]::Now.ToString('yyyy.dd.MM HH:mm:ss'))" 
            $filename=((Split-Path ($MyInvocation.MyCommand.Path))+"lastrunlog.txt")
            if( Test-Path $filename -PathType Leaf ) {
                $c = Get-Content $filename
                $cl = $c -split '`n'
                $s = ((@($s) + $cl) | select -First 200)
            }
            Out-File -InputObject ($s -join "`r`n") -FilePath $filename
        }
    } catch {
          Write-EventLog -Source "SharePoint Warmup Script"  -Category 0 -ComputerName "." -EntryType Error -LogName "Application" `
            -Message "SharePoint Warmup failed for url ""$($url)""." -EventId 1001

        $s = "Last run failed for url ""$($url)"": $([DateTime]::Now.ToString('yyyy.dd.MM HH:mm:ss')) : $($_.Exception.Message)" 
        $filename=((Split-Path ($MyInvocation.MyCommand.Path))+"lastrunlog.txt")
        if( Test-Path $filename -PathType Leaf ) {
          $c = Get-Content $filename
          $cl = $c -split '`n'
          $s = ((@($s) + $cl) | select -First 200)
        }
        Out-File -InputObject ($s -join "`r`n") -FilePath $filename
    }
}
Function CrawlSitemap
{
    Param(
        [parameter(Mandatory=$true)]
        [string] $SiteMapUrl
    );

    $SiteMapXml = Invoke-WebRequest -Uri $SiteMapUrl -UseBasicParsing -TimeoutSec 180;
    $Urls = ([xml]$SiteMapXml).urlset.ChildNodes
    ForEach ($Url in $Urls){
        $Loc = $Url.loc;
        try{
            $result = Invoke-WebRequest -Uri $Loc -UseBasicParsing -TimeoutSec 180;
            Write-Host $result.StatusCode - $Loc;
        }catch [System.Net.WebException] {
            Write-Warning (([int]$_.Exception.Response.StatusCode).ToString() + " - " + $Loc);
        }
    }
}
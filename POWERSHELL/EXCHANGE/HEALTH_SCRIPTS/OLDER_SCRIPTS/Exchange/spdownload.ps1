$SharePointClientDll = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\SharePoint Client Components\'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Location') + "ISAPI\Microsoft.SharePoint.Client.dll")
Add-Type -Path $SharePointClientDll 


function DownloadFileFromOneDrive{
	param (
	        $DownloadURL = "$( throw 'DownloadURL is a mandatory Parameter' )",
			$PSCredentials = "$( throw 'credentials is a mandatory Parameter' )",
			$DownloadPath  = "$( throw 'DownloadPath is a mandatory Parameter' )"
		  )
	process{
		$DownloadURI = New-Object System.Uri($DownloadURL);
		$SharepointHost = "https://" + $DownloadURI.Host
		$soCredentials =  New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($PSCredentials.UserName.ToString(),$PSCredentials.password) 
		$clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($SharepointHost)
		$clientContext.Credentials = $soCredentials;
		$destFile = $DownloadPath + [System.IO.Path]::GetFileName($DownloadURI.LocalPath)
		$fileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($clientContext, $DownloadURI.LocalPath);
		$fstream = New-Object System.IO.FileStream($destFile, [System.IO.FileMode]::Create);
		$fileInfo.Stream.CopyTo($fstream)
		$fstream.Flush()
		$fstream.Close()
		Write-Host ("File downloaded to " + ($destFile))
	}
}

$cred = Get-Credential
#exmample use
DownloadFileFromOneDrive -DownloadURL $args[0] -PSCredentials $cred -DownloadPath 'c:\Temp\'
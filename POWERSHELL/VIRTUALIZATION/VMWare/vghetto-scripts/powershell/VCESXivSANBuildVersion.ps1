﻿Function Get-VCVersion {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function extracts the vCenter Server (Windows or VCSA) build from your env
        and maps it to https://kb.vmware.com/kb/2143838 to retrieve the version and release date
    .EXAMPLE
        Get-VCVersion
#>
    param(
        [Parameter(Mandatory=$false)][VMware.VimAutomation.ViCore.Util10.VersionedObjectImpl]$Server
    )

    # Pulled from https://kb.vmware.com/kb/2143838
    $vcenterBuildVersionMappings = @{
        "5973321"="vCenter 6.5 Update 1,2017-07-27"
        "5705665"="vCenter 6.5 0e Express Patch 3,2017-06-15"
        "5318154"="vCenter 6.5 0d Express Patch 2,2017-04-18"
        "5318200"="vCenter 6.0 Update 3b,2017-04-13"
        "5183549"="vCenter 6.0 Update 3a,2017-03-21"
        "5112527"="vCenter 6.0 Update 3,2017-02-24"
        "4541947"="vCenter 6.0 Update 2a,2016-11-22"
        "3634793"="vCenter 6.0 Update 2,2016-03-16"
        "3339083"="vCenter 6.0 Update 1b,2016-01-07"
        "3018524"="vCenter 6.0 Update 1,2015-09-10"
        "2776511"="vCenter 6.0.0b,2015-07-07"
        "2656760"="vCenter 6.0.0a,2015-04-16"
        "2559268"="vCenter 6.0 GA,2015-03-12"
        "4180647"="vCenter 5.5 Update 3e,2016-08-04"
        "3721164"="vCenter 5.5 Update 3d,2016-04-14"
        "3660016"="vCenter 5.5 Update 3c,2016-03-29"
        "3252642"="vCenter 5.5 Update 3b,2015-12-08"
        "3142196"="vCenter 5.5 Update 3a,2015-10-22"
        "3000241"="vCenter 5.5 Update 3,2015-09-16"
        "2646482"="vCenter 5.5 Update 2e,2015-04-16"
        "2001466"="vCenter 5.5 Update 2,2014-09-09"
        "1945274"="vCenter 5.5 Update 1c,2014-07-22"
        "1891313"="vCenter 5.5 Update 1b,2014-06-12"
        "1750787"="vCenter 5.5 Update 1a,2014-04-19"
        "1750596"="vCenter 5.5.0c,2014-04-19"
        "1623099"="vCenter 5.5 Update 1,2014-03-11"
        "1378903"="vCenter 5.5.0a,2013-10-31"
        "1312299"="vCenter 5.5 GA,2013-09-22"
        "3900744"="vCenter 5.1 Update 3d,2016-05-19"
        "3070521"="vCenter 5.1 Update 3b,2015-10-01"
        "2669725"="vCenter 5.1 Update 3a,2015-04-30"
        "2207772"="vCenter 5.1 Update 2c,2014-10-30"
        "1473063"="vCenter 5.1 Update 2,2014-01-16"
        "1364037"="vCenter 5.1 Update 1c,2013-10-17"
        "1235232"="vCenter 5.1 Update 1b,2013-08-01"
        "1064983"="vCenter 5.1 Update 1,2013-04-25"
        "880146"="vCenter 5.1.0a,2012-10-25"
        "799731"="vCenter 5.1 GA,2012-09-10"
        "3891028"="vCenter 5.0 U3g,2016-06-14"
        "3073236"="vCenter 5.0 U3e,2015-10-01"
        "2656067"="vCenter 5.0 U3d,2015-04-30"
        "1300600"="vCenter 5.0 U3,2013-10-17"
        "913577"="vCenter 5.0 U2,2012-12-20"
        "755629"="vCenter 5.0 U1a,2012-07-12"
        "623373"="vCenter 5.0 U1,2012-03-15"
        "5318112"="vCenter 6.5.0c Express Patch 1b,2017-04-13"
        "5178943"="vCenter 6.5.0b,2017-03-14"
        "4944578"="vCenter 6.5.0a Express Patch 01,2017-02-02"
        "4602587"="vCenter 6.5,2016-11-15"
        "5326079"="vCenter 6.0 Update 3b,2017-04-13"
        "5183552"="vCenter 6.0 Update 3a,2017-03-21"
        "5112529"="vCenter 6.0 Update 3,2017-02-24"
        "4541948"="vCenter 6.0 Update 2a,2016-11-22"
        "4191365"="vCenter 6.0 Update 2m,2016-09-15"
        "3634794"="vCenter 6.0 Update 2,2016-03-15"
        "3339084"="vCenter 6.0 Update 1b,2016-01-07"
        "3018523"="vCenter 6.0 Update 1,2015-09-10"
        "2776510"="vCenter 6.0.0b,2015-07-07"
        "2656761"="vCenter 6.0.0a,2015-04-16"
        "2559267"="vCenter 6.0 GA,2015-03-12"
        "4180648"="vCenter 5.5 Update 3e,2016-08-04"
        "3730881"="vCenter 5.5 Update 3d,2016-04-14"
        "3660015"="vCenter 5.5 Update 3c,2016-03-29"
        "3255668"="vCenter 5.5 Update 3b,2015-12-08"
        "3154314"="vCenter 5.5 Update 3a,2015-10-22"
        "3000347"="vCenter 5.5 Update 3,2015-09-16"
        "2646489"="vCenter 5.5 Update 2e,2015-04-16"
        "2442329"="vCenter 5.5 Update 2d,2015-01-27"
        "2183111"="vCenter 5.5 Update 2b,2014-10-09"
        "2063318"="vCenter 5.5 Update 2,2014-09-09"
        "1623101"="vCenter 5.5 Update 1,2014-03-11"
        "1476327"="vCenter 5.5.0b,2013-12-22"
        "1398495"="vCenter 5.5.0a,2013-10-31"
        "1312298"="vCenter 5.5 GA,2013-09-22"
        "3868380"="vCenter 5.1 Update 3d,2016-05-19"
        "3630963"="vCenter 5.1 Update 3c,2016-03-29"
        "3072314"="vCenter 5.1 Update 3b,2015-10-01"
        "2306353"="vCenter 5.1 Update 3,2014-12-04"
        "1882349"="vCenter 5.1 Update 2a,2014-07-01"
        "1474364"="vCenter 5.1 Update 2,2014-01-16"
        "1364042"="vCenter 5.1 Update 1c,2013-10-17"
        "1123961"="vCenter 5.1 Update 1a,2013-05-22"
        "1065184"="vCenter 5.1 Update 1,2013-04-25"
        "947673"="vCenter 5.1.0b,2012-12-20"
        "880472"="vCenter 5.1.0a,2012-10-25"
        "799730"="vCenter 5.1 GA,2012-08-13"
        "3891027"="vCenter 5.0 U3g,2016-06-14"
        "3073237"="vCenter 5.0 U3e,2015-10-01"
        "2656066"="vCenter 5.0 U3d,2015-04-30"
        "2210222"="vCenter 5.0 U3c,2014-11-20"
        "1917469"="vCenter 5.0 U3a,2014-07-01"
        "1302764"="vCenter 5.0 U3,2013-10-17"
        "920217"="vCenter 5.0 U2,2012-12-20"
        "804277"="vCenter 5.0 U1b,2012-08-16"
        "759855"="vCenter 5.0 U1a,2012-07-12"
        "455964"="vCenter 5.0 GA,2011-08-24"
    }

    if(-not $Server) {
        $Server = $global:DefaultVIServer
    }

    $vcBuildNumber = $Server.Build
    $vcName = $Server.Name
    $vcOS = $Server.ExtensionData.Content.About.OsType
    $vcVersion,$vcRelDate = "Unknown","Unknown"

    if($vcenterBuildVersionMappings.ContainsKey($vcBuildNumber)) {
        ($vcVersion,$vcRelDate) = $vcenterBuildVersionMappings[$vcBuildNumber].split(",")
    }

    $tmp = [pscustomobject] @{
        Name = $vcName;
        Build = $vcBuildNumber;
        Version = $vcVersion;
        OS = $vcOS;
        ReleaseDate = $vcRelDate;
    }
    $tmp
}

Function Get-ESXiVersion {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function extracts the ESXi build from your env and maps it to
        https://kb.vmware.com/kb/2143832 to extract the version and release date
    .PARAMETER ClusterName
        Name of the vSphere Cluster to retrieve ESXi version information
    .EXAMPLE
        Get-ESXiVersion -ClusterName VSAN-Cluster
#>
    param(
        [Parameter(Mandatory=$true)][String]$ClusterName
    )

    # Pulled from https://kb.vmware.com/kb/2143832
    $esxiBuildVersionMappings = @{
        "5969303"="ESXi 6.5 U1,2017-07-27"
        "5310538"="ESXi 6.5.0d,2017-04-18"
        "5224529"="ESXi 6.5 Express Patch 1a,2017-03-28"
        "5146846"="ESXi 6.5 Patch 01,2017-03-09"
        "4887370"="ESXi 6.5.0a,2017-02-02"
        "4564106"="ESXi 6.5 GA,2016-11-15"
        "5572656"="ESXi 6.0 Patch 5,2017-06-06"
        "5251623"="ESXi 6.0 Express Patch 7c,2017-03-28"
        "5224934"="ESXi 6.0 Express Patch 7a,2017-03-28"
        "5050593"="ESXi 6.0 Update 3,2017-02-24"
        "4600944"="ESXi 6.0 Patch 4,2016-11-22"
        "4510822"="ESXi 6.0 Express Patch 7,2016-10-17"
        "4192238"="ESXi 6.0 Patch 3,2016-08-04"
        "3825889"="ESXi 6.0 Express Patch 6,2016-05-12"
        "3620759"="ESXi 6.0 Update 2,2016-03-16"
        "3568940"="ESXi 6.0 Express Patch 5,2016-02-23"
        "3380124"="ESXi 6.0 Update 1b,2016-01-07"
        "3247720"="ESXi 6.0 Express Patch 4,2015-11-25"
        "3073146"="ESXi 6.0 U1a Express Patch 3,2015-10-06"
        "3029758"="ESXi 6.0 U1,2015-09-10"
        "2809209"="ESXi 6.0.0b,2015-07-07"
        "2715440"="ESXi 6.0 Express Patch 2,2015-05-14"
        "2615704"="ESXi 6.0 Express Patch 1,2015-04-09"
        "2494585"="ESXi 6.0 GA,2015-03-12"
        "5230635"="ESXi 5.5 Express Patch 11,2017-03-28"
        "4722766"="ESXi 5.5 Patch 10,2016-12-20"
        "4345813"="ESXi 5.5 Patch 9,2016-09-15"
        "4179633"="ESXi 5.5 Patch 8,2016-08-04"
        "3568722"="ESXi 5.5 Express Patch 10,2016-02-22"
        "3343343"="ESXi 5.5 Express Patch 9,2016-01-04"
        "3248547"="ESXi 5.5 Update 3b,2015-12-08"
        "3116895"="ESXi 5.5 Update 3a,2015-10-06"
        "3029944"="ESXi 5.5 Update 3,2015-09-16"
        "2718055"="ESXi 5.5 Patch 5,2015-05-08"
        "2638301"="ESXi 5.5 Express Patch 7,2015-04-07"
        "2456374"="ESXi 5.5 Express Patch 6,2015-02-05"
        "2403361"="ESXi 5.5 Patch 4,2015-01-27"
        "2302651"="ESXi 5.5 Express Patch 5,2014-12-02"
        "2143827"="ESXi 5.5 Patch 3,2014-10-15"
        "2068190"="ESXi 5.5 Update 2,2014-09-09"
        "1892794"="ESXi 5.5 Patch 2,2014-07-01"
        "1881737"="ESXi 5.5 Express Patch 4,2014-06-11"
        "1746018"="ESXi 5.5 Update 1a,2014-04-19"
        "1746974"="ESXi 5.5 Express Patch 3,2014-04-19"
        "1623387"="ESXi 5.5 Update 1,2014-03-11"
        "1474528"="ESXi 5.5 Patch 1,2013-12-22"
        "1331820"="ESXi 5.5 GA,2013-09-22"
        "3872664"="ESXi 5.1 Patch 9,2016-05-24"
        "3070626"="ESXi 5.1 Patch 8,2015-10-01"
        "2583090"="ESXi 5.1 Patch 7,2015-03-26"
        "2323236"="ESXi 5.1 Update 3,2014-12-04"
        "2191751"="ESXi 5.1 Patch 6,2014-10-30"
        "2000251"="ESXi 5.1 Patch 5,2014-07-31"
        "1900470"="ESXi 5.1 Express Patch 5,2014-06-17"
        "1743533"="ESXi 5.1 Patch 4,2014-04-29"
        "1612806"="ESXi 5.1 Express Patch 4,2014-02-27"
        "1483097"="ESXi 5.1 Update 2,2014-01-16"
        "1312873"="ESXi 5.1 Patch 3,2013-10-17"
        "1157734"="ESXi 5.1 Patch 2,2013-07-25"
        "1117900"="ESXi 5.1 Express Patch 3,2013-05-23"
        "1065491"="ESXi 5.1 Update 1,2013-04-25"
        "1021289"="ESXi 5.1 Express Patch 2,2013-03-07"
        "914609"="ESXi 5.1 Patch 1,2012-12-20"
        "838463"="ESXi 5.1.0a,2012-10-25"
        "799733"="ESXi 5.1.0 GA,2012-09-10"
        "3982828"="ESXi 5.0 Patch 13,2016-06-14"
        "3086167"="ESXi 5.0 Patch 12,2015-10-01"
        "2509828"="ESXi 5.0 Patch 11,2015-02-24"
        "2312428"="ESXi 5.0 Patch 10,2014-12-04"
        "2000308"="ESXi 5.0 Patch 9,2014-08-28"
        "1918656"="ESXi 5.0 Express Patch 6,2014-07-01"
        "1851670"="ESXi 5.0 Patch 8,2014-05-29"
        "1489271"="ESXi 5.0 Patch 7,2014-01-23"
        "1311175"="ESXi 5.0 Update 3,2013-10-17"
        "1254542"="ESXi 5.0 Patch 6,2013-08-29"
        "1117897"="ESXi 5.0 Express Patch 5,2013-05-15"
        "1024429"="ESXi 5.0 Patch 5,2013-03-28"
        "914586"="ESXi 5.0 Update 2,2012-12-20"
        "821926"="ESXi 5.0 Patch 4,2012-09-27"
        "768111"="ESXi 5.0 Patch 3,2012-07-12"
        "721882"="ESXi 5.0 Express Patch 4,2012-06-14"
        "702118"="ESXi 5.0 Express Patch 3,2012-05-03"
        "653509"="ESXi 5.0 Express Patch 2,2012-04-12"
        "623860"="ESXi 5.0 Update 1,2012-03-15"
        "515841"="ESXi 5.0 Patch 2,2011-12-15"
        "504890"="ESXi 5.0 Express Patch 1,2011-11-03"
        "474610"="ESXi 5.0 Patch 1,2011-09-13"
        "469512"="ESXi 5.0 GA,2011-08-24"
    }

    $cluster = Get-Cluster -Name $ClusterName -ErrorAction SilentlyContinue
    if($cluster -eq $null) {
        Write-Host -ForegroundColor Red "Error: Unable to find vSAN Cluster $ClusterName ..."
        break
    }

    $results = @()
    foreach ($vmhost in $cluster.ExtensionData.Host) {
        $vmhost_view = Get-View $vmhost -Property Name, Config, ConfigManager.ImageConfigManager

        $esxiName = $vmhost_view.name
        $esxiBuild = $vmhost_view.Config.Product.Build
        $esxiVersionNumber = $vmhost_view.Config.Product.Version
        $esxiVersion,$esxiRelDate,$esxiOrigInstallDate = "Unknown","Unknown","N/A"

        if($esxiBuildVersionMappings.ContainsKey($esxiBuild)) {
            ($esxiVersion,$esxiRelDate) = $esxiBuildVersionMappings[$esxiBuild].split(",")
        }

        # Install Date API was only added in 6.5
        if($esxiVersionNumber -eq "6.5.0") {
            $imageMgr = Get-View $vmhost_view.ConfigManager.ImageConfigManager
            $esxiOrigInstallDate = $imageMgr.installDate()
        }

        $tmp = [pscustomobject] @{
            Name = $esxiName;
            Build = $esxiBuild;
            Version = $esxiVersion;
            ReleaseDate = $esxiRelDate;
            OriginalInstallDate = $esxiOrigInstallDate;
        }
        $results+=$tmp
    }
    $results
}

Function Get-VSANVersion {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function extracts the ESXi build from your env and maps it to
        https://kb.vmware.com/kb/2150753 to extract the vSAN version and release date
    .PARAMETER ClusterName
        Name of a vSAN Cluster to retrieve vSAN version information
    .EXAMPLE
        Get-VSANVersion -ClusterName VSAN-Cluster
#>
    param(
        [Parameter(Mandatory=$true)][String]$ClusterName
    )

    # Pulled from https://kb.vmware.com/kb/2150753
    $vsanBuildVersionMappings = @{
        "5969303"="vSAN 6.6.1,ESXi 6.5 Update 1,2017-07-27"
        "5310538"="vSAN 6.6,ESXi 6.5.0d,2017-04-18"
        "5224529"="vSAN 6.5 Express Patch 1a,ESXi 6.5 Express Patch 1a,2017-03-28"
        "5146846"="vSAN 6.5 Patch 01,ESXi 6.5 Patch 01,2017-03-09"
        "4887370"="vSAN 6.5.0a,ESXi 6.5.0a,2017-02-02"
        "4564106"="vSAN 6.5,ESXi 6.5 GA,2016-11-15"
        "5572656"="vSAN 6.2 Patch 5,ESXi 6.0 Patch 5,2017-06-06"
        "5251623"="vSAN 6.2 Express Patch 7c,ESXi 6.0 Express Patch 7c,2017-03-28"
        "5224934"="vSAN 6.2 Express Patch 7a,ESXi 6.0 Express Patch 7a,2017-03-28"
        "5050593"="vSAN 6.2 Update 3,ESXi 6.0 Update 3,2017-02-24"
        "4600944"="vSAN 6.2 Patch 4,ESXi 6.0 Patch 4,2016-11-22"
        "4510822"="vSAN 6.2 Express Patch 7,ESXi 6.0 Express Patch 7,2016-10-17"
        "4192238"="vSAN 6.2 Patch 3,ESXi 6.0 Patch 3,2016-08-04"
        "3825889"="vSAN 6.2 Express Patch 6,ESXi 6.0 Express Patch 6,2016-05-12"
        "3620759"="vSAN 6.2,ESXi 6.0 Update 2,2016-03-16"
        "3568940"="vSAN 6.1 Express Patch 5,ESXi 6.0 Express Patch 5,2016-02-23"
        "3380124"="vSAN 6.1 Update 1b,ESXi 6.0 Update 1b,2016-01-07"
        "3247720"="vSAN 6.1 Express Patch 4,ESXi 6.0 Express Patch 4,2015-11-25"
        "3073146"="vSAN 6.1 U1a (Express Patch 3),ESXi 6.0 U1a (Express Patch 3),2015-10-06"
        "3029758"="vSAN 6.1,ESXi 6.0 U1,2015-09-10"
        "2809209"="vSAN 6.0.0b,ESXi 6.0.0b,2015-07-07"
        "2715440"="vSAN 6.0 Express Patch 2,ESXi 6.0 Express Patch 2,2015-05-14"
        "2615704"="vSAN 6.0 Express Patch 1,ESXi 6.0 Express Patch 1,2015-04-09"
        "2494585"="vSAN 6.0,ESXi 6.0 GA,2015-03-12"
        "5230635"="vSAN 5.5 Express Patch 11,ESXi 5.5 Express Patch 11,2017-03-28"
        "4722766"="vSAN 5.5 Patch 10,ESXi 5.5 Patch 10,2016-12-20"
        "4345813"="vSAN 5.5 Patch 9,ESXi 5.5 Patch 9,2016-09-15"
        "4179633"="vSAN 5.5 Patch 8,ESXi 5.5 Patch 8,2016-08-04"
        "3568722"="vSAN 5.5 Express Patch 10,ESXi 5.5 Express Patch 10,2016-02-22"
        "3343343"="vSAN 5.5 Express Patch 9,ESXi 5.5 Express Patch 9,2016-01-04"
        "3248547"="vSAN 5.5 Update 3b,ESXi 5.5 Update 3b,2015-12-08"
        "3116895"="vSAN 5.5 Update 3a,ESXi 5.5 Update 3a,2015-10-06"
        "3029944"="vSAN 5.5 Update 3,ESXi 5.5 Update 3,2015-09-16"
        "2718055"="vSAN 5.5 Patch 5,ESXi 5.5 Patch 5,2015-05-08"
        "2638301"="vSAN 5.5 Express Patch 7,ESXi 5.5 Express Patch 7,2015-04-07"
        "2456374"="vSAN 5.5 Express Patch 6,ESXi 5.5 Express Patch 6,2015-02-05"
        "2403361"="vSAN 5.5 Patch 4,ESXi 5.5 Patch 4,2015-01-27"
        "2302651"="vSAN 5.5 Express Patch 5,ESXi 5.5 Express Patch 5,2014-12-02"
        "2143827"="vSAN 5.5 Patch 3,ESXi 5.5 Patch 3,2014-10-15"
        "2068190"="vSAN 5.5 Update 2,ESXi 5.5 Update 2,2014-09-09"
        "1892794"="vSAN 5.5 Patch 2,ESXi 5.5 Patch 2,2014-07-01"
        "1881737"="vSAN 5.5 Express Patch 4,ESXi 5.5 Express Patch 4,2014-06-11"
        "1746018"="vSAN 5.5 Update 1a,ESXi 5.5 Update 1a,2014-04-19"
        "1746974"="vSAN 5.5 Express Patch 3,ESXi 5.5 Express Patch 3,2014-04-19"
        "1623387"="vSAN 5.5,ESXi 5.5 Update 1,2014-03-11"
    }

    $cluster = Get-Cluster -Name $ClusterName -ErrorAction SilentlyContinue
    if($cluster -eq $null) {
        Write-Host -ForegroundColor Red "Error: Unable to find vSAN Cluster $ClusterName ..."
        break
    }

    $results = @()
    foreach ($vmhost in $cluster.ExtensionData.Host) {
        $vmhost_view = Get-View $vmhost -Property Name, Config, ConfigManager.ImageConfigManager

        $esxiName = $vmhost_view.name
        $esxiBuild = $vmhost_view.Config.Product.Build
        $esxiVersionNumber = $vmhost_view.Config.Product.Version
        $vsanVersion,$esxiVersion,$esxiRelDate = "Unknown","Unknown","Unknown"

        # Technically as of vSAN 6.2 Mgmt API, this information is already built in natively within
        # the product to retrieve ESXi/VC/vSAN Versions
        # See https://github.com/lamw/vghetto-scripts/blob/master/powershell/VSANVersion.ps1
        if($vsanBuildVersionMappings.ContainsKey($esxiBuild)) {
            ($vsanVersion,$esxiVersion,$esxiRelDate) = $vsanBuildVersionMappings[$esxiBuild].split(",")
        }

        $tmp = [pscustomobject] @{
            Name = $esxiName;
            Build = $esxiBuild;
            VSANVersion = $vsanVersion;
            ESXiVersion = $esxiVersion;
            ReleaseDate = $esxiRelDate;
        }
        $results+=$tmp
    }
    $results
}
# NAME: Get-ActiveSyncDeviceInfo.ps1 
# 
# AUTHOR: Jan Egil Ring 
# EMAIL: jan.egil.ring@powershell.no 
# 
# COMMENT: Script to retrieve all ActiveSync-devices registered within the Exchange-organization. 
#          A conversion-table for Apple-devices are provided, you might remove this if you want to 
#          retrieve the real DeviceUserAgent-names for those devices. 
#          The script outputs objects to make it easier working with the results, i.e. to export 
#          the output using Export-Csv, sort them, group them and so on. 
#          Works with both Exchange 2007 and Exchange 2010. Since a new cmdlet, Get-ActiveSyncDevice, 
#          exist in Exchange 2010, you might want to use that when working against Exchange 2010. 
#           
#          For more information, see the following blog-post: 
#          http://blog.powershell.no/2010/09/26/getting-an-overview-of-all-activesync-devices-in-the-exchange-organization 
#       
# You have a royalty-free right to use, modify, reproduce, and 
# distribute this script file in any way you find useful, provided that 
# you agree that the creator, owner above has no warranty, obligations, 
# or liability for such use. 
# 
# VERSION HISTORY: 
# 1.0 26.09.2010 - Initial release 
# 
# Note: Updated with table information from http://exchangescripts.blogspot.com/2012/03/exchange-activesync-device-info.html
#       and bastardized further by Zachary Loeber.
# 
 
#Loop through each mailbox
Function Custom-Function
{ 
    [CmdletBinding()] 
    PARAM([Parameter(Position=0, Mandatory=$true, ValueFromPipeline = $true)] [PSObject[]]$InputObject) 
 
    BEGIN
    { 
        $Mailboxes=@()
        $ActiveSyncInfo = @()
    } 
    PROCESS
    { 
        $Mailboxes += $InputObject
        Foreach ($mailbox in $Mailboxes) 
        { 
            $devices = Get-ActiveSyncDeviceStatistics -Mailbox $mailbox.Identity | Select-Object DeviceType,DevicePolicyApplied,LastSuccessSync,DeviceUserAgent 
             
            #If the current mailbox has an ActiveSync device associated, loop through each device 
            if ($devices) 
            { 
                foreach ($device in $devices)
                { 
                    #Conversion table for Apple-devices 
                    switch ($device.DeviceUserAgent) 
                    { 
                        "Apple-(null)/704.11" {$DeviceUserAgent = "iPhone"}  
                         "Apple-iPad/702.367" {$DeviceUserAgent = "iPad"}  
                         "Apple-iPad/702.405" {$DeviceUserAgent = "iPad"}  
                         "Apple-iPad/702.500" {$DeviceUserAgent = "iPad"}  
                         "Apple-iPad1C1/803.134" {$DeviceUserAgent = "iPad"}  
                         "Apple-iPad1C1/803.148" {$DeviceUserAgent = "iPad"}  
                         "Apple-iPad1C1/806.190" {$DeviceUserAgent = "iPad"}  
                         "Apple-iPad1C1/807.4" {$DeviceUserAgent = "iPad"}  
                         "Apple-iPad1C1/808.7" {$DeviceUserAgent = "iPad"}  
                         "Apple-iPad1C1/810.3" {$DeviceUserAgent = "iPad"}  
                         "Apple-iPad1C1/811.2" {$DeviceUserAgent = "iPad"}  
                         "Apple-iPad1C1/812.1" {$DeviceUserAgent = "iPad"}  
                         "Apple-iPad1C1/901.334" {$DeviceUserAgent = "iPad"}  
                         "Apple-iPad1C1/901.405" {$DeviceUserAgent = "iPad"}  
                         "Apple-iPad1C1/901.528800004" {$DeviceUserAgent = "iPad"}  
                         "Apple-iPad2C1/806.191" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C1/807.4" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C1/808.7" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C1/810.2" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C1/811.2" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C1/812.1" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C1/901.334" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C1/901.405" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C1/901.524800004" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C1/901.527400004" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C2/806.191" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C2/807.4" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C2/808.7" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C2/810.2" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C2/811.2" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C2/812.1" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C2/901.334" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C2/901.405" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C3/806.191" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C3/807.4" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C3/808.8" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C3/810.2" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C3/811.2" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C3/812.1" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C3/901.334" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPad2C3/901.405" {$DeviceUserAgent = "iPad2"}  
                         "Apple-iPhone/501.347" {$DeviceUserAgent = "iPhone"}  
                         "Apple-iPhone/502.108" {$DeviceUserAgent = "iPhone"}  
                         "Apple-iPhone/503.1" {$DeviceUserAgent = "iPhone"}  
                         "Apple-iPhone/506.136" {$DeviceUserAgent = "iPhone"}  
                         "Apple-iPhone/507.77" {$DeviceUserAgent = "iPhone"}  
                         "Apple-iPhone/508.11" {$DeviceUserAgent = "iPhone"}  
                         "Apple-iPhone/701.341" {$DeviceUserAgent = "iPhone"}  
                         "Apple-iPhone/701.400" {$DeviceUserAgent = "iPhone"}  
                         "Apple-iPhone/703.144" {$DeviceUserAgent = "iPhone"}  
                         "Apple-iPhone/704.11" {$DeviceUserAgent = "iPhone"}  
                         "Apple-iPhone/705.18" {$DeviceUserAgent = "iPhone"}  
                         "DTG-iPhone/4.0" {$DeviceUserAgent = "iPhone"}  
                         "Apple-iPhone1C2/801.293" {$DeviceUserAgent = "iPhone 3G"}  
                         "Apple-iPhone1C2/801.306" {$DeviceUserAgent = "iPhone 3G"}  
                         "Apple-iPhone1C2/801.400" {$DeviceUserAgent = "iPhone 3G"}  
                         "Apple-iPhone1C2/802.117" {$DeviceUserAgent = "iPhone 3G"}  
                         "Apple-iPhone1C2/803.148" {$DeviceUserAgent = "iPhone 3G"}  
                         "Apple-iPhone2C1/801.23000013" {$DeviceUserAgent = "iPhone 3GS"}  
                         "Apple-iPhone2C1/801.27400002" {$DeviceUserAgent = "iPhone 3GS"}  
                         "Apple-iPhone2C1/801.293" {$DeviceUserAgent = "iPhone 3GS"}  
                         "Apple-iPhone2C1/801.306" {$DeviceUserAgent = "iPhone 3GS"}  
                         "Apple-iPhone2C1/801.400" {$DeviceUserAgent = "iPhone 3GS"}  
                         "Apple-iPhone2C1/802.117" {$DeviceUserAgent = "iPhone 3GS"}  
                         "Apple-iPhone2C1/803.14800001" {$DeviceUserAgent = "iPhone 3GS"}  
                         "Apple-iPhone2C1/806.190" {$DeviceUserAgent = "iPhone 3GS"}  
                         "Apple-iPhone2C1/807.4" {$DeviceUserAgent = "iPhone 3GS"}  
                         "Apple-iPhone2C1/808.7" {$DeviceUserAgent = "iPhone 3GS"}  
                         "Apple-iPhone2C1/810.2" {$DeviceUserAgent = "iPhone 3GS"}  
                         "Apple-iPhone2C1/811.2" {$DeviceUserAgent = "iPhone 3GS"}  
                         "Apple-iPhone2C1/812.1" {$DeviceUserAgent = "iPhone 3GS"}  
                         "Apple-iPhone2C1/901.334" {$DeviceUserAgent = "iPhone 3GS"}  
                         "Apple-iPhone2C1/901.405" {$DeviceUserAgent = "iPhone 3GS"}  
                         "Apple-iPhone2C1/901.531300005" {$DeviceUserAgent = "iPhone 3GS"}  
                         "Apple-iPhone3C1/801.293" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C1/801.306" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C1/801.400" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C1/802.117" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C1/803.148" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C1/806.190" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C1/807.4" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C1/808.7" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C1/810.2" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C1/811.2" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C1/812.1" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C1/901.334" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C1/901.405" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C1/901.522000016" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C1/901.530200002" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C1/901.531300005" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C3/805.128" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C3/805.200" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C3/805.303" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C3/805.401" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C3/805.501" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C3/805.600" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C3/901.334" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone3C3/901.405" {$DeviceUserAgent = "iPhone 4"}  
                         "Apple-iPhone4C1/901.334" {$DeviceUserAgent = "iPhone 4S"}  
                         "Apple-iPhone4C1/901.405" {$DeviceUserAgent = "iPhone 4S"}  
                         "Apple-iPhone4C1/901.406" {$DeviceUserAgent = "iPhone 4S"}  
                         "Apple-iPhone4C1/902.512700003" {$DeviceUserAgent = "iPhone 4S"}  
                         "Apple-iPod/506.137" {$DeviceUserAgent = "iPod"}  
                         "Apple-iPod/506.138" {$DeviceUserAgent = "iPod"}  
                         "Apple-iPod/507.77" {$DeviceUserAgent = "iPod"}  
                         "Apple-iPod/507.7700001" {$DeviceUserAgent = "iPod"}  
                         "Apple-iPod/508.11" {$DeviceUserAgent = "iPod"}  
                         "Apple-iPod/508.1100001" {$DeviceUserAgent = "iPod"}  
                         "Apple-iPod/701.341" {$DeviceUserAgent = "iPod"}  
                         "Apple-iPod/703.144" {$DeviceUserAgent = "iPod"}  
                         "Apple-iPod/703.145" {$DeviceUserAgent = "iPod"}  
                         "Apple-iPod/703.146" {$DeviceUserAgent = "iPod"}  
                         "Apple-iPod/704.11" {$DeviceUserAgent = "iPod"}  
                         "Apple-iPod/705.18" {$DeviceUserAgent = "iPod"}  
                         "Apple-iPod2C1/801.293" {$DeviceUserAgent = "iPod 2G"}  
                         "Apple-iPod2C1/801.400" {$DeviceUserAgent = "iPod 2G"}  
                         "Apple-iPod2C1/802.117" {$DeviceUserAgent = "iPod 2G"}  
                         "Apple-iPod2C1/803.148" {$DeviceUserAgent = "iPod 2G"}  
                         "Apple-iPod3C1/801.293" {$DeviceUserAgent = "iPod 3G"}  
                         "Apple-iPod3C1/801.400" {$DeviceUserAgent = "iPod 3G"}  
                         "Apple-iPod3C1/802.117" {$DeviceUserAgent = "iPod 3G"}  
                         "Apple-iPod3C1/803.148" {$DeviceUserAgent = "iPod 3G"}  
                         "Apple-iPod3C1/806.190" {$DeviceUserAgent = "iPod 3G"}  
                         "Apple-iPod3C1/807.4" {$DeviceUserAgent = "iPod 3G"}  
                         "Apple-iPod3C1/808.7" {$DeviceUserAgent = "iPod 3G"}  
                         "Apple-iPod3C1/810.2" {$DeviceUserAgent = "iPod 3G"}  
                         "Apple-iPod3C1/811.2" {$DeviceUserAgent = "iPod 3G"}  
                         "Apple-iPod3C1/812.1" {$DeviceUserAgent = "iPod 3G"}  
                         "Apple-iPod3C1/901.334" {$DeviceUserAgent = "iPod 3G"}  
                         "Apple-iPod3C1/901.405" {$DeviceUserAgent = "iPod 3G"}  
                         "Apple-iPod4C1/802.117" {$DeviceUserAgent = "iPod 4G"}  
                         "Apple-iPod4C1/802.118" {$DeviceUserAgent = "iPod 4G"}  
                         "Apple-iPod4C1/803.148" {$DeviceUserAgent = "iPod 4G"}  
                         "Apple-iPod4C1/806.190" {$DeviceUserAgent = "iPod 4G"}  
                         "Apple-iPod4C1/807.4" {$DeviceUserAgent = "iPod 4G"}  
                         "Apple-iPod4C1/808.7" {$DeviceUserAgent = "iPod 4G"}  
                         "Apple-iPod4C1/810.2" {$DeviceUserAgent = "iPod 4G"}  
                         "Apple-iPod4C1/811.2" {$DeviceUserAgent = "iPod 4G"}  
                         "Apple-iPod4C1/812.1" {$DeviceUserAgent = "iPod 4G"}  
                         "Apple-iPod4C1/901.334" {$DeviceUserAgent = "iPod 4G"}  
                         "Apple-iPod4C1/901.405" {$DeviceUserAgent = "iPod 4G"}  
                         "Android-EAS/0.1" {$DeviceUserAgent = "Android"}  
                         "Android/0.3" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.0.2119383124.172283" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/2.0" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.0.2117352549.117459" {$DeviceUserAgent = "Android"}  
                         "Android/4.0.2-EAS-1.3" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/2.1.2115213027.63843" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.0.2115272322.67241" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/3.10.000.037285.405" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.0.2121132845.194526" {$DeviceUserAgent = "Android"}  
                         "Android/3.2.1-EAS-1.2" {$DeviceUserAgent = "Android"}  
                         "Android/4.0.3-EAS-1.3" {$DeviceUserAgent = "Android"}  
                         "Android/3.1-EAS-1.2" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/3.10.000.082258.502" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.0.2115402813.76204" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/1.0.2116383032.94711" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/3.10.000.074499.651" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.0.2120413252.187846" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.0.2117372546.119931" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/3.10.000.095282.531" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/3.10.000.206244.605" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.5.2120232624.181895" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.0.2114392227.55840" {$DeviceUserAgent = "Android"}  
                         "Android/3.2-EAS-1.2" {$DeviceUserAgent = "Android"}  
                         "Android" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/1.0.2116333166.92077" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.0.2119182913.157066" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.0.2115272323.67241" {$DeviceUserAgent = "Android"}  
                         "Android/3.0-EAS-1.2" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.0.2116123015.76346" {$DeviceUserAgent = "Android"}  
                         "Android/3.0.1-EAS-1.2" {$DeviceUserAgent = "Android"}  
                         "Android/4.0.1-EAS-1.3" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.0.2121203215.199782" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/3.1.2115372068.74619" {$DeviceUserAgent = "Android"}  
                         "Android/4.0.4-EAS-1.3" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/3.10.000.178661.605" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/3.10.000.259408.651" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/2.2.2120372569.189233" {$DeviceUserAgent = "Android"}  
                         "Android/3.2.2-EAS-1.2" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/3.10.000.147855.605" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.0.2122332740.238988" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/3.1.2116372767.93900" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.5.2118263164.139119" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.5.2122332410.238988" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.5.2120152733.177334" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.0.2119182466.156684" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.0.2122302045.226132" {$DeviceUserAgent = "Android"}  
                         "Android/3.2.1 KRAKD by jcarrz1-EAS-1.2" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.0.2122362126.238988" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.0.2119322552.167079" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/3.10.000.141744.573" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/3.10.000.065587.651" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.5.2118372724.147551" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/3.10" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.5.2118302654.142550" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.5.2119393325.173702" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.0.2117152551.98788" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/3.1.2115222666.64656" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.5.2120242026.182442" {$DeviceUserAgent = "Android"}  
                         "Android/IceCreamSandwich-EAS-1.3" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.5.2211302633.275410" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/3.1.2118362635.147199" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/4.0.2118382031.147471" {$DeviceUserAgent = "Android"}  
                         "Android-EAS/3.10.000.209954.531" {$DeviceUserAgent = "Android"}  
                         "Android/0.3/14.0" {$DeviceUserAgent = "Android"}  
                         "HTC-HTCAmaze4G(4D6F7869SAM)/2.1500" {$DeviceUserAgent = "Android"}  
                         "HTC-HTCGlacier(4D6F7869SAM)/2.1404" {$DeviceUserAgent = "Android"}  
                         "HTC-T-MobileG2(4D6F7869SAM)/2.1402" {$DeviceUserAgent = "Android"}  
                         "HTC-T-MobileG26B101ABFSAM/2.1400" {$DeviceUserAgent = "Android"}  
                         "Moto-Blur/3.0" {$DeviceUserAgent = "Android"}  
                         "Moto-DROID BIONIC/5.5.1" {$DeviceUserAgent = "Android"}  
                         "Moto-DROID PRO/" {$DeviceUserAgent = "Android"}  
                         "Moto-DROID PRO/3.4.2" {$DeviceUserAgent = "Android"}  
                         "Moto-DROID Pro/4.5.1" {$DeviceUserAgent = "Android"}  
                         "Moto-DROID RAZR/6.5.1" {$DeviceUserAgent = "Android"}  
                         "Moto-DROID X2/4.4.1" {$DeviceUserAgent = "Android"}  
                         "Moto-DROID X2/4.5.1" {$DeviceUserAgent = "Android"}  
                         "Moto-DROID2 GLOBAL/4.5.1" {$DeviceUserAgent = "Android"}  
                         "Moto-DROID2/4.5.1" {$DeviceUserAgent = "Android"}  
                         "Moto-DROID3/5.5.1" {$DeviceUserAgent = "Android"}  
                         "Moto-DROID4/6.5.1" {$DeviceUserAgent = "Android"}  
                         "Moto-DROIDX/4.5.1" {$DeviceUserAgent = "Android"}  
                         "Moto-MB508/3.4.2" {$DeviceUserAgent = "Android"}  
                         "Moto-MB520/3.4.2" {$DeviceUserAgent = "Android"}  
                         "Moto-MB525/3.4.2" {$DeviceUserAgent = "Android"}  
                         "Moto-MB611/2.5" {$DeviceUserAgent = "Android"}  
                         "Moto-MB611/4.0" {$DeviceUserAgent = "Android"}  
                         "Moto-MB855/4.5.1" {$DeviceUserAgent = "Android"}  
                         "Moto-MB860/4.5.141" {$DeviceUserAgent = "Android"}  
                         "Moto-MB860/4.5.91" {$DeviceUserAgent = "Android"}  
                         "Moto-MB865/5.5.1" {$DeviceUserAgent = "Android"}  
                         "Moto-Milestone X/4.5.1" {$DeviceUserAgent = "Android"}  
                         "Moto-Milestone X2/45.0.25" {$DeviceUserAgent = "Android"}  
                         "Moto-Morrison/1.0" {$DeviceUserAgent = "Android"}  
                         "Moto-Motorola Electrify/4.5.1" {$DeviceUserAgent = "Android"}  
                         "TouchDown(MSRPC)/7.0.0012" {$DeviceUserAgent = "Android"}  
                         "SAMSUNG-GT-I9000/100.20304" {$DeviceUserAgent = "Android"}  
                         "SAMSUNG-SAMSUNG-SGH-I997/100.20303" {$DeviceUserAgent = "Android"}  
                         "TouchDown(MSRPC)/6.4.0007" {$DeviceUserAgent = "Android"}  
                         "MSFT-SPhone/5.2.1108" {$DeviceUserAgent = "Android"}  
                         "Microsoft-PocketPC/3.0" {$DeviceUserAgent = "Android"}  
                         "SAMSUNG-SAMSUNG-SGH-I897/100.20303" {$DeviceUserAgent = "Android"}  
                         "Apache-HttpClient/UNAVAILABLE (java 1.4)" {$DeviceUserAgent = "Android"}  
                         "Apache-HttpClient/UNAVAILABLE (java 1.4)" {$DeviceUserAgent = "Android"}  
                         "EASClient/1.0" {$DeviceUserAgent = "Android"}  
                         "EASClient/2.1300" {$DeviceUserAgent = "Android"}  
                         "Moxier-Android/1.0" {$DeviceUserAgent = "Android"}  
                         "Remoba-cdma_droid2-VZW-DROID2" {$DeviceUserAgent = "Android"}  
                         "Remoba-inc-FRF91-ADR6300" {$DeviceUserAgent = "Android"}  
                         "Remoba-olympus-4.5.91-MB860" {$DeviceUserAgent = "Android"}  
                         "Remoba-SGH-I727-GINGERBREAD-SAMSUNG-SGH-I727" {$DeviceUserAgent = "Android"}  
                         "Remoba-SGH-I897-FROYO-SAMSUNG-SGH-I897" {$DeviceUserAgent = "Android"}  
                         "Remoba-SGH-T589-FROYO-SGH-T589" {$DeviceUserAgent = "Android"}  
                         "Remoba-SPH-D710-GINGERBREAD-SPH-D710" {$DeviceUserAgent = "Android"}  
                         "RoadSync-Android/2.503" {$DeviceUserAgent = "Android"}  
                         "RoadSync-Android/2.503" {$DeviceUserAgent = "Android"}  
                         "RoadSync-Android/2.503" {$DeviceUserAgent = "Android"}  
                         "RoadSync-Android/2.503" {$DeviceUserAgent = "Android"}  
                         "SAMSUNG-GT-I9000/100.202" {$DeviceUserAgent = "Android"}  
                         "SAMSUNG-GT-P7510/100.301" {$DeviceUserAgent = "Android"}  
                         "SAMSUNG-SAMSUNG-SGH-I897/100.202" {$DeviceUserAgent = "Android"}  
                         "SAMSUNG-SAMSUNG-SGH-I997/100.202" {$DeviceUserAgent = "Android"}  
                         "SAMSUNGSCHI405/2.3.5-EAS-1.2" {$DeviceUserAgent = "Android"}  
                         "SAMSUNG-SCH-M828C/100.202" {$DeviceUserAgent = "Android"}  
                         "SAMSUNGSGHI896/0.3" {$DeviceUserAgent = "Android"}  
                         "SAMSUNG-SGH-I897/100.202" {$DeviceUserAgent = "Android"}  
                         "SAMSUNG-SGH-T499/100.202" {$DeviceUserAgent = "Android"}  
                         "SAMSUNG-SGH-T589/100.202" {$DeviceUserAgent = "Android"}  
                         "SAMSUNG-SGH-T839/100.202" {$DeviceUserAgent = "Android"}  
                         "SAMSUNG-SGH-T959V/100.202" {$DeviceUserAgent = "Android"}  
                         "SAMSUNG-SHW-M110S/100.202" {$DeviceUserAgent = "Android"}  
                         "SAMSUNGSPHD700/100.202" {$DeviceUserAgent = "Android"}  
                         "SAMSUNG-SPH-M580/100.202" {$DeviceUserAgent = "Android"}  
                         "SAMSUNG-SPH-M820-BST/100.202" {$DeviceUserAgent = "Android"}  
                         "SAMSUNGSPHP100/0.3" {$DeviceUserAgent = "Android"}  
                         "TouchDown(MSRPC)" {$DeviceUserAgent = "Android"}  
                         "TouchDown(MSRPC)/5.1.0026" {$DeviceUserAgent = "Android"}  
                         "TouchDown(MSRPC)/5.1.0028" {$DeviceUserAgent = "Android"}  
                         "TouchDown(MSRPC)/6.0.0002" {$DeviceUserAgent = "Android"}  
                         "TouchDown(MSRPC)/6.1.0007" {$DeviceUserAgent = "Android"}  
                         "TouchDown(MSRPC)/6.1.0010" {$DeviceUserAgent = "Android"}  
                         "TouchDown(MSRPC)/6.2.0012" {$DeviceUserAgent = "Android"}  
                         "TouchDown(MSRPC)/6.5.0002" {$DeviceUserAgent = "Android"}  
                         "TouchDown(MSRPC)/7.0.0012" {$DeviceUserAgent = "Android"}  
                         "TouchDown(MSRPC)/7.1.00012/" {$DeviceUserAgent = "Android"}  
                         "RoadSync-Android/2.502" {$DeviceUserAgent = "Android"}  
                         "RoadSync-Android/2.503" {$DeviceUserAgent = "Android"}  
                         "RoadSync-Android/2.503" {$DeviceUserAgent = "Android"}  
                         "US670/1.0" {$DeviceUserAgent = "Android"}  
                         "US760/1.0" {$DeviceUserAgent = "Android"}  
                         "VM670/0.3" {$DeviceUserAgent = "Android"}  
                         "Vortex/1.0" {$DeviceUserAgent = "Android"}  
                         "VS700/1.0" {$DeviceUserAgent = "Android"}  
                         "LS670/0.3" {$DeviceUserAgent = "Android"}  
                         "LW690/1.0" {$DeviceUserAgent = "Android"}  
                         "Ally/1.0" {$DeviceUserAgent = "Android"}  
                         "GARM-A50/1.6" {$DeviceUserAgent = "Android"}  
                         "GARM-A50/2.1-update1" {$DeviceUserAgent = "Android"}  
                         "LGMC-LGEAS/3.81" {$DeviceUserAgen = "Android"}  
                         "LGMC-LGEAS/4.13" {$DeviceUserAgen = "Android"}  
                         "LGMC-LGEAS/4.161V" {$DeviceUserAgen = "Android"}  
                         "LGMC-LGEAS/4.23" {$DeviceUserAgen = "Android"}  
                         "LGMC-LGEAS/4.38IV" {$DeviceUserAgen = "Android"}  
                         "PANTECH-PantechP9070/1.0" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-GalaxyNexus(4D6F7869SAM)/2.1501" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-GT-I9100/100.20303" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-GT-P7510/100.302" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-GT-S5360L/100.20306" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-GT-S5830/100.20305" {$DeviceUserAgen = "Android"}  
                         "SAMSUNGSAMSUNGSGH/100" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SAMSUNG-SGH-I777/100.20304" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SAMSUNG-SGH-I777/100.20306" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SAMSUNG-SGH-I897(4D6F7869SAM)/2.1501" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SAMSUNG-SGH-I897/100.20305" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SGH-I727/100.20305" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SGH-I727/100.20306" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SGH-I997/100.20306" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SGH-T679/100.20305" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SGH-T759/100.20303" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SGH-T859/100.302" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SGH-T959V/100.20305" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SGH-T959V/100.20306" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SGH-T989(4D6F7869SAM)/2.1501" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SGH-T989/100.20305" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SGH-T989/100.20306" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SHW-M110S/100.20303" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SPH-D600/100.20304" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SPH-D700/100.20303" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SPH-D700/100.20304" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SPH-D700/100.20305" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SPH-D700/100.20306" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SPH-D710(4D6F7869SAM)/2.1500" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SPH-D710/100.20304" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SPH-D710/100.20306" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-YP-G1/100.20305" {$DeviceUserAgen = "Android"}  
                         "TouchDown(MSRPC)/6.4.0002" {$DeviceUserAgen = "Android"}  
                         "TouchDown(MSRPC)/7.1.0005" {$DeviceUserAgen = "Android"}  
                         "VM701/1.0" {$DeviceUserAgen = "Android"}  
                         "Moto-XT603/5.5.1" {$DeviceUserAgen = "Android"}  
                         "MyS1Device2" {$DeviceUserAgen = "Android"}  
                         "SAMSUNG-SCH-I905/100.302" {$DeviceUserAgent = "Android"}  
                         "SAMSUNG-SGH-I957/100.302" {$DeviceUserAgent = "Android"}  
                         "WindowsMail/16.2.3237.0215" {$DeviceUserAgent = "Windows 8"}  
                         "NokiaE73/2.02(0)MailforExchange 3gpp-gba" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaN97/2.09(208)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaN97/3.00(73)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaE71x/2.09(208)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaN958GB/3.00(50)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaN800/3.00(0)MailforExchange 3gpp-gba" {$DeviceUserAgent = "Nokia-EAS"}  
                         "Nokia5530/2.09(208)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaE71x/1.00(0)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaE71x/2.09(158)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "Nokia5230/2.09(206)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaE71/1.0" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaE600/1.00(0)MailforExchange 3gpp-gba" {$DeviceUserAgent = "Nokia-EAS"}  
                         "Nokia6730c/2.02(0)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaE752/2.01(0)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaE71/2.09(158)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "Nokia6790s1b/2.09(160)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaN800/1.00(0)MailforExchange 3gpp-gba" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaE71/3.00(50)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaE722/2.01(0)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "Nokia5800XpressMusic/2.07(37)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "Nokia5800XpressMusic/2.09(158)MailforExchange 3gpp-gba" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaE63/2.09(158)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaC700/1.00(0)MailforExchange 3gpp-gba" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaN97/2.09(206)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "Nokia5230c/3.00(73)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaN95/2.09(158)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "Nokia5800XpressMusic/2.09(210)MailforExchange 3gpp-gba" {$DeviceUserAgent = "Nokia-EAS"}  
                         "Nokia5800XpressMusic/2.09(188)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "Nokia5530/2.09(188)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaC600/2.05(0)MailforExchange 3gpp-gba" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaE90/2.05(0)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaE63/2.07(22)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaE721/2.02(0)MailforExchange 3gpp-gba" {$DeviceUserAgent = "Nokia-EAS"}  
                         "NokiaE722/2.02(0)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "Nokia5800XpressMusic/2.09(206)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "Nokia5800XpressMusic/2.09(162)MailforExchange" {$DeviceUserAgent = "Nokia-EAS"}  
                         "N900/1.1" {$DeviceUserAgent = "Nokia-EAS"}  
                         "MeeGo-MfE/1.2" {$DeviceUserAgent = "MeeGo"}  
                         "Palm/1.0.1" {$DeviceUserAgent = "WebOS"}  
                         "RIM-Playbook/2.0" {$DeviceUserAgent = "WebOS"}  
                         "MSFT-PPC/5.2.0" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1000" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1001" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1001" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1004" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1005" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1005" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1203" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1203" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1203" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1206" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1207" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1207" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1208" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1303" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1303" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1303" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1403" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1404" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1405" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1406" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1501" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1501" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1604" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1605" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1605" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1607" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.1608" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.400" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.402" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.402" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.402" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.404" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.404" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.404" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.408" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5026" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5028" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5063" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5063" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5063" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5063" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5070" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5081" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5081" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5082" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5083" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5083" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5084" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5086" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5086" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5086" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5087" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5087" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5087" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5087" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5089" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5094" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5095" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5306" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5309" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5309" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5312" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5312" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5500" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.5500" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.2.700" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.1000" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.1001" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.1001" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.1004" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.1108" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.1110" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.1150" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.1152" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.1308" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.1310" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.1602" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.1603" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.1606" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.201" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.202" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.301" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.304" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.402" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.403" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.501" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.5027" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.5080" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.5080" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.603" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.604" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/4.0" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.1.2201" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.1.2202" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.1.2301" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.1.2302" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.1.3301" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-PPC/5.1.3502" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.1.2300" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.1.2400" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.1.3002" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.301" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-SPhone/5.2.402" {$DeviceUserAgent = "Windows Mobile"}  
                         "Microsoft-PocketPC/3.0" {$DeviceUserAgent = "Windows Mobile"}  
                         "MSFT-WP/7.0.7003" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.0.7004" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.0.7004" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.0.7004" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.0.7008" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.0.7008" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.0.7389" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.0.7390" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.0.7390" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.0.7390" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.0.7392" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.0.7392" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.0.7392" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.0.7403" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.10.7720" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.10.7720" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.10.7720" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.10.7720" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.10.7740" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.10.7740" {$DeviceUserAgent = "Windows Phone 7"}  
                         "MSFT-WP/7.10.8107" {$DeviceUserAgent = "Windows Phone 7"}  
                         "Mozilla/4.0 (compatible; ZuneHD 4.5)" {$DeviceUserAgent = "Microsoft Zune"}  
                         "PalmOne-TreoAce/1.02m01" {$DeviceUserAgent = "Palm OS"}  
                         "PalmOne-TreoAce/1.52" {$DeviceUserAgent = "Palm OS"}  
                         "PalmOne-TreoAce/1.53" {$DeviceUserAgent = "Palm OS"}  
                         "PalmOne-TreoAce/2.011m01" {$DeviceUserAgent = "Palm OS"}  
                         "PalmOne-TreoAce/2.01m01" {$DeviceUserAgent = "Palm OS"}  
                         "RoadSync-S60/5.0" {$DeviceUserAgent = "Symbian"}  
                         "RoadSync-S60/4.0" {$DeviceUserAgent = "Symbian"}  
                         "RoadSync/3.0" {$DeviceUserAgent = "Symbian"}  
                         "FastSync" {$DeviceUserAgent = "Helio Ocean"}  
                         "BlackBerry/5.2.200 UNTRUSTED/1.0" {$DeviceUserAgent = "Blackberry"}  
                        default 
                        {
                            $DeviceUserAgent = $device.DeviceUserAgent
                        } 
                    } 
                     
                    #Create a new object and add custom note properties for each device 
                    $deviceobj = New-Object -TypeName psobject 
                    $deviceobj | Add-Member -Name User -Value $mailbox.DisplayName -MemberType NoteProperty 
                    $deviceobj | Add-Member -Name DeviceType -Value $device.DeviceType -MemberType NoteProperty 
                    $deviceobj | Add-Member -Name DeviceUserAgent -Value $DeviceUserAgent -MemberType NoteProperty 
                    $deviceobj | Add-Member -Name DevicePolicyApplied -Value $device.DevicePolicyApplied -MemberType NoteProperty 
                    $deviceobj | Add-Member -Name LastSuccessSync -Value ($device.LastSuccessSync).ToShortDateString() -MemberType NoteProperty 
                     
                    #Write the custom object to the pipeline 
                    $ActiveSyncInfo += $deviceobj 
                } 
            } 
        }
        $ActiveSyncInfo | Export-Csv ".\ActiveSyncDevices.csv" -NoTypeInformation
    }
}
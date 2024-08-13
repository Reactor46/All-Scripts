﻿---
title: Troubleshooting OWA.Proxy Health Set
TOCTitle: Troubleshooting OWA.Proxy Health Set
ms:assetid: 1eaa26ad-b489-402a-ad2d-bfae3b083f42
ms:mtpsurl: https://technet.microsoft.com/en-us/library/ms.exch.scom.owa.proxy(v=EXCHG.150)
ms:contentKeyID: 49720734
ms.date: 10/08/2015
mtps_version: v=EXCHG.150
---

<div data-xmlns="http://www.w3.org/1999/xhtml">

<div class="topic" data-xmlns="http://www.w3.org/1999/xhtml" data-msxsl="urn:schemas-microsoft-com:xslt" data-cs="http://msdn.microsoft.com/en-us/">

<div data-asp="http://msdn2.microsoft.com/asp">

# Troubleshooting OWA.Proxy Health Set

</div>

<div id="mainSection">

<div id="mainBody">

<span> </span>

_**Applies to:** Exchange Server 2013_

_**Topic Last Modified:** 2015-03-09_

The OWA.Proxy health set monitors the overall health of the Outlook Web App service.

If you receive an alert that specifies that the OWA.Proxy is unhealthy, this indicates an issue that may prevent users from using the Outlook Web App service.

<span id="EXP"></span>

<div>

## Explanation

The Outlook Web App service is monitored by using the following probes and monitors.


<table>
<colgroup>
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
<col style="width: 25%" />
</colgroup>
<thead>
<tr class="header">
<th>Probe</th>
<th>Health Set</th>
<th>Dependencies</th>
<th>Associated Monitors</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p>OWAProxyTestProbe</p></td>
<td><p>OWA.Proxy</p></td>
<td><p>Active Directory</p></td>
<td><p>OWAProxyTestMonitor</p></td>
</tr>
<tr class="even">
<td><p>OWAanonymouscalendarProbe</p></td>
<td><p>OWA.Proxy</p></td>
<td><p>Active Directory</p></td>
<td><p>OWAProxyTestMonitor</p></td>
</tr>
</tbody>
</table>


For more information about probes and monitors, see [Server health and performance](https://technet.microsoft.com/en-us/library/jj150551\(v=exchg.150\)).

</div>

<div>

## Common issues

This probe may fail for any of the following common reasons:

  - The application pool that’s hosted on the monitored Client Access server (CAS) is not working correctly.

  - The monitoring account credentials are incorrect.

  - The Domain Controllers are not responding.

</div>

<div>

## User Action

It's possible that the service recovered after it issued the alert. Therefore, when you receive an alert that specifies that the health set is unhealthy, first verify that the issue still exists. If the issue does exist, perform the appropriate recovery actions outlined in the following sections.

<span id="verify"></span>

<div>

## Verifying the issue still exists

1.  Identify the health set name and the server name in the alert.

2.  The message details provide information about the exact cause of the alert. In most cases, the message details provide sufficient troubleshooting information to identify the root cause. If the message details are not clear, do the following:
    
    1.  Open the Exchange Management Shell, and then run the following command to retrieve the details of the health set that issued the alert:
        
            Get-ServerHealth <server name> | ?{$_.HealthSetName -eq "<health set name>"}
        
        For example, to retrieve the OWA.Proxy health set details about server1.contoso.com, run the following command:
        
            Get-ServerHealth server1.contoso.com | ?{$_.HealthSetName -eq "OWA.Proxy"}
    
    2.  Review the command output to determine which monitor reported the error. The **AlertValue** value for the monitor that issued the alert will be `Unhealthy`.
    
    3.  Rerun the associated probe for the monitor that’s in an unhealthy state. Refer to the table in the Explanation section to find the associated probe. To do this, run the following command:
        
            Invoke-MonitoringProbe <health set name>\<probe name> -Server <server name> | Format-List
        
        For example, assume that the failing monitor is **OWAProxyTestMonitor**. The probe associated with that monitor is **OWAProxyTestProbe**. To run that probe on server1.contoso.com, run the following command:
        
            Invoke-MonitoringProbe OWA.Proxy\OWAProxyTestProbe -Server server1.contoso.com | Format-List
    
    4.  In the command output, review the **Result** value of the probe. If the value is **Succeeded**, the issue was a transient error, and it no longer exists. Otherwise, refer to the recovery steps outlined in the following sections.

</div>

<span id="TestMonitors"></span>

<div>

## OWAProxyTestMonitor Recovery Actions

When you receive an alert from a health set, the email message contains the following information:

  - Name of the CAS that sent the alert

  - Full exception trace of the last error, including diagnostic data and specific HTTP header information  
    
    **Note**   You can use the information in the full exception trace to help troubleshoot the issue.

  - Time and date when the issue occurred

To troubleshoot this issue, follow these steps:

1.  Review the protocol logs on the CA servers. Protocol logs are located in the *\<exchange server installation directory\>*\\Logging\\HttpProxy*\\\<protocol\>* folder on the CAS.

2.  Create a test user account, and then log on to the CAS by using the test user account. For example, log on by using: https:// *\<servername\>*/owa.

3.  Start IIS Manager, and then connect to the server that’s reporting the issue to determine whether the **MSExchangeServicesAppPools** application pools are running on the CAS.

4.  Click **Application Pools**, and then recycle the **MSExchangeOWAAppPool** and **MSExchangeOWACalendarAppPool** application pools by running the following command from the Shell:
    
        %SystemRoot%\System32\inetsrv\Appcmd recycle APPPOOL MSExchangeOWAAppPool
    
        %SystemRoot%\System32\inetsrv\Appcmd recycle APPPOOL MSExchangeOWACalendarAppPool

5.  Rerun the associated probe as shown in step 2c in the Verifying the issue still exists section.

6.  If the issue still exists, recycle the IIS service by using the IISReset utility.

7.  Rerun the associated probe as shown in step 2c in the Verifying the issue still exists section.

8.  If the issue still exists, restart the server.

9.  After the server restarts, rerun the associated probe as shown in step 2c in the Verifying the issue still exists section.

10. If the probe continues to fail, you may need assistance to resolve this issue. Contact a Microsoft Support professional to resolve this issue. To contact a Microsoft Support professional, visit the [Exchange Server Solutions Center](http://go.microsoft.com/fwlink/p/?linkid=180809). In the navigation pane, click **Support options and resources** and use one of the options listed under **Get technical support** to contact a Microsoft Support professional. Because your organization may have a specific procedure for directly contacting Microsoft Product Support Services, be sure to review your organization's guidelines first.

</div>

</div>

<div>

## For More Information

[What's new in Exchange 2013](https://technet.microsoft.com/en-us/library/jj150540\(v=exchg.150\))

[Exchange 2013 cmdlets](https://technet.microsoft.com/en-us/library/bb124413\(v=exchg.150\))

[Managing Outlook Web App](https://technet.microsoft.com/en-us/3814b665-01e8-4881-9a44-163f14789ee4\(exchg.150\)#managing)

</div>

</div>

<span> </span>

</div>

</div>

</div>


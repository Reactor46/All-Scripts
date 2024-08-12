﻿---
title: Troubleshooting RPS.Proxy Health Set
TOCTitle: Troubleshooting RPS.Proxy Health Set
ms:assetid: a5058323-5d86-438a-ad4a-fa4292310e98
ms:mtpsurl: https://technet.microsoft.com/en-us/library/ms.exch.scom.rps.proxy(v=EXCHG.150)
ms:contentKeyID: 49720858
ms.date: 10/08/2015
mtps_version: v=EXCHG.150
---

<div data-xmlns="http://www.w3.org/1999/xhtml">

<div class="topic" data-xmlns="http://www.w3.org/1999/xhtml" data-msxsl="urn:schemas-microsoft-com:xslt" data-cs="http://msdn.microsoft.com/en-us/">

<div data-asp="http://msdn2.microsoft.com/asp">

# Troubleshooting RPS.Proxy Health Set

</div>

<div id="mainSection">

<div id="mainBody">

<span> </span>

_**Applies to:** Exchange Server 2013_

_**Topic Last Modified:** 2015-03-09_

The RPS.Proxy health set monitors the overall health of the Remote PowerShell service.

If you receive an alert specifying that the RPS.Proxy is unhealthy, this indicates an issue that might prevent you from using Remote PowerShell to access Exchange.

<span id="EXP"></span>

<div>

## Explanation

The RPS service is monitored using the following probes and monitors:


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
<td><p>RPSProxyTestProbe</p></td>
<td><p>RPS.Proxy</p></td>
<td><p>Active Directory</p></td>
<td><p>RPSProxyTestMonitor</p></td>
</tr>
</tbody>
</table>


For more information about probes and monitors, see [Server health and performance](https://technet.microsoft.com/en-us/library/jj150551\(v=exchg.150\)).

</div>

<div>

## Common issues

When this probe fails there can be multiple reasons for the problem. Some of the more common issues include the following:

  - The application pool that is hosted on the monitored CAS server is not working properly.

  - The monitoring account credentials are incorrect.

  - The Domain Controllers are not responding.

</div>

<div>

## User Action

It is possible that the service was able to recover after issuing the alert. Therefore, when you receive an alert that specifies that the health set is unhealthy, the first thing you should do is to verify that the issue still exists. If it does, then perform the appropriate recovery actions outlined in the following sections.

<span id="verify"></span>

<div>

## Verifying the issue still exists

1.  Identify the health set name and the server name in the alert.

2.  The message details provide information about what exactly caused the alert to be raised. In most cases, the message details would provide sufficient troubleshooting information to identify the root cause. If the message details are not clear, do the following:
    
    1.  Open Exchange Management Shell and run the following command to retrieve the details of the health set that issued the alert:
        
            Get-ServerHealth <server name> | ?{$_.HealthSetName -eq "<health set name>"}
        
        For example, to retrieve the RPS.Proxy health set details on server1.contoso.com run the following command:
        
            Get-ServerHealth server1.contoso.com | ?{$_.HealthSetName -eq "RPS.Proxy"}
    
    2.  Review the command output and determine the monitor that reported the error. The **AlertValue** for the monitor that issued the alert will read `Unhealthy`.
    
    3.  Rerun the associated probe for the monitor that is in unhealthy state. Refer to the table in the Explanation section to find the associated probe. To do so, run the following command:
        
            Invoke-MonitoringProbe <health set name>\<probe name> -Server <server name> | Format-List
        
        For example, let's assume that the failing monitor was **RPSProxyTestMonitor**. The probe associated with that monitor is **RPSProxyTestProbe**. To run that probe on server server1.contoso.com, run the following command:
        
            Invoke-MonitoringProbe RPS.Proxy\RPSProxyTestProbe -Server server1.contoso.com | Format-List
    
    4.  In the command output, review the **Result** of the probe. If the value is **Succeeded**, then the issue was a transient error and no longer exists. Otherwise refer to the recovery steps outlined in the following sections.

</div>

<span id="TestMonitors"></span>

<div>

## RPSProxyTestMonitor Recovery Actions

When receiving an alert from a health set, the email will contain the following information:

  - The name of the CAS server that sent the alert.

  - Full exception trace including eroor messages, diagnostic data and specific HTTP header information. The information in the full exception trace can be used to help troubleshoot the issue.

  - The time and date when the issue occurred.

To help troubleshoot this issue, perform the following:

1.  Review the protocol logs on CAS servers. Protocol logs are located in the *\<exchange server installation directory\>*\\Logging\\HttpProxy*\\\<protocol\>* folder on the CAS server.

2.  Create a test user account, and then logon to the CAS server by using the test user account, for example https:// *\<servername\>*/owa

3.  Start IIS Manager and connect to the server that is reporting the issue and verify that the MSExchangePowerShellFrontEndAppPool is running on CAS server.

4.  Click on **Application Pools**, and then recycle the **MSExchangePowerShellFrontEndAppPool** application pool by running the following command from the Exchange Management Shell:
    
        %SystemRoot%\System32\inetsrv\Appcmd recycle MSExchangePowerShellFrontEndAppPool

5.  Rerun the associated probe as shown in step 2.c. in the Verifying the issue still exists section.

6.  If the issue still exists, recycle the IIS service using the IISReset utility.

7.  Rerun the associated probe as shown in step 2.c. in the Verifying the issue still exists section.

8.  If the issue still exists, restart the server.

9.  After the server restarts, rerun the associated probe as shown in step 2.c. in the Verifying the issue still exists section.

10. If the probe continues to fail, you may need assistance in resolving this issue. Contact a Microsoft Support professional to resolve this issue. To contact a Microsoft Support professional, visit the [Exchange Server Solutions Center](http://go.microsoft.com/fwlink/p/?linkid=180809). In the navigation pane, click **Support options and resources** and use one of the options listed under **Get technical support** to contact a Microsoft Support professional. Because your organization may have a specific procedure for directly contacting Microsoft Product Support Services, be sure to review your organization's guidelines first.

</div>

</div>

<div>

## For More Information

[What's new in Exchange 2013](https://technet.microsoft.com/en-us/library/jj150540\(v=exchg.150\))

[Exchange 2013 cmdlets](https://technet.microsoft.com/en-us/library/bb124413\(v=exchg.150\))

</div>

</div>

<span> </span>

</div>

</div>

</div>


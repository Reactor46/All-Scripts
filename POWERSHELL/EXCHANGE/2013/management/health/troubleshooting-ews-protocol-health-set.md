﻿---
title: Troubleshooting EWS.Protocol Health Set
TOCTitle: Troubleshooting EWS.Protocol Health Set
ms:assetid: 826b2d5b-adbb-4bf5-94b6-0a8de2e3aac0
ms:mtpsurl: https://technet.microsoft.com/en-us/library/ms.exch.scom.ews.protocol(v=EXCHG.150)
ms:contentKeyID: 49720829
ms.date: 10/08/2015
mtps_version: v=EXCHG.150
---

<div data-xmlns="http://www.w3.org/1999/xhtml">

<div class="topic" data-xmlns="http://www.w3.org/1999/xhtml" data-msxsl="urn:schemas-microsoft-com:xslt" data-cs="http://msdn.microsoft.com/en-us/">

<div data-asp="http://msdn2.microsoft.com/asp">

# Troubleshooting EWS.Protocol Health Set

</div>

<div id="mainSection">

<div id="mainBody">

<span> </span>

_**Applies to:** Exchange Server 2013, Project Server 2013_

_**Topic Last Modified:** 2015-03-09_

The EWS.Protocol health set monitors the Exchange Web Services (EWS) communications protocol on the Mailbox server. The EWS.Protocol health set is closely related to the following health sets:

[Troubleshooting EWS Health Set](troubleshooting-ews-health-set.md)

[Troubleshooting EWS.Proxy Health Set](troubleshooting-ews-proxy-health-set.md)

If you receive an alert that specifies that the EWS.Protocol is unhealthy, this indicates an issue that may prevent your users from accessing Exchange.

<span id="EXP"></span>

<div>

## Explanation

The EWS.Protocol health set is composed of the following probes:

1.  EwsSelfTestProbe

2.  EwsDeepTestProbe

The EwsSelfTestProbe does not depend on the Information Store. However, the EwsDeepTestProbe probe depends on the Information Store. Both of these probes perform EWS operations on the Mailbox server, and they use the same authentication method as a Client Access server (CAS). EwsSeftTestProbe calls the **ConvertId** method, and EwsDeepTestProbe calls the **GetFolder** method.


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
<td><p>EwsSelfTestProbe</p></td>
<td><p>EWS.Protocol</p></td>
<td><p>Active Directory</p></td>
<td><p>EWSSelfTestMonitor</p></td>
</tr>
<tr class="even">
<td><p>EwsDeepTestProbe</p></td>
<td><p>EWS.Protocol</p></td>
<td><p>Information Store</p></td>
<td><p>EWSDeepTestMonitor</p></td>
</tr>
</tbody>
</table>


For more information about probes and monitors, see [Server health and performance](https://technet.microsoft.com/en-us/library/jj150551\(v=exchg.150\)).

When you receive an alert from this HealthSet, the email message contains the following information:

1.  Name of the Mailbox server on which the alert originated

2.  Full exception trace of the last error, including diagnostic data and specific HTTP headers information

3.  Time when the incident occurred

</div>

<div>

## Common issues

This probe can fail for any of the following common reasons:

  - The EWS Application pool on the monitored Mailbox server is not functioning correctly.

  - The Domain Controllers are not responding, or they cannot communicate with the Mailbox server.

  - The user’s database is not mounted, or the Information Store is unavailable for a specific mailbox.

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
        
        For example, to retrieve the EWS.Protocol health set details about server1.contoso.com, run the following command:
        
            Get-ServerHealth server1.contoso.com | ?{$_.HealthSetName -eq "EWS.Protocol"}
    
    2.  Review the command output to determine which monitor reported the error. The **AlertValue** value for the monitor that issued the alert will be `Unhealthy`.
    
    3.  Rerun the associated probe for the monitor that is in an unhealthy state. Refer to the table in the Explanation section to find the associated probe. To do this, run the following command:
        
            Invoke-MonitoringProbe <health set name>\<probe name> -Server <server name> | Format-List
        
        For example, assume that the failing monitor is **EWSSelfTestMonitor**. The probe associated with that monitor is **EWSSelfTestProbe**. To run that probe on server1.contoso.com, run the following command:
        
            Invoke-MonitoringProbe EWS.Protocol\EWSSelfTestProbe -Server server1.contoso.com | Format-List
    
    4.  In the command output, review the **Result** value of the probe. If the value is **Succeeded**, the issue was a transient error, and it no longer exists. Otherwise, refer to the recovery steps outlined in the following sections.

</div>

<span id="TestMonitors"></span>

<div>

## EWSSelfTestMonitor and EWSDeepTestMonitor Recovery Actions

This monitor alert is typically issued for Mailbox servers.

1.  Start IIS Manager, and then connect to the server that’s reporting the issue to determine whether the MSExchangeServicesAppPool is running on both CA and Mailbox servers.

2.  Locate the MailboxDatabase for the failed probes, and then verify that the MailboxDatabase is active for the MailboxServer, and that the Information Store is healthy.

3.  Click **Application Pools**, and then recycle the **MSExchangeServicesAppPool** application pool by running the following command from the Shell:
    
        %SystemRoot%\System32\inetsrv\Appcmd recycle MSExchangeServicesAppPool

4.  Rerun the associated probe as shown in step 2c in the Verifying the issue still exists section.

5.  If the issue still exists, restart the IIS service by using the IISReset utility.

6.  Rerun the associated probe as shown in step 2c in the Verifying the issue still exists section.

7.  If the issue still exits, review the protocol log files on the Mailbox server. On the Mailbox server, the logs reside in the *\<exchange server installation directory\>*\\Logging\\Ews folder.

8.  Create a test user account, and then log on by using the test user account against the given Mailbox server on port 444 https://*\<servername\>*:444/ews/exchange.asmx. If the test is successful, an issue may affect the specific mailbox database or Mailbox server on which the monitoring mailbox is located. Try to repeat this step by using a test account on that database.

9.  Check for any alerts on the EWS.Protocol Health Set that might indicate a problem that affects the specific Mailbox server.

10. If the issue still exists, restart the server. To do this, first failover the databases hosted on the server by using the following command:
    
        Set-MailboxServer server1.contoso.com -DatabaseCopyActivationDisabledAndMoveNow $true
    
    In this and all subsequent code examples, replace *server1.contoso.com* with the actual server name.

11. Verify that all the databases have been moved off the server that is reporting the issue. To do this, run the following command:
    
        Get-MailboxDatabaseCopyStatus -Server server1.contoso.com | Group Status
    
    If the command output shows no active copies on the server, the server is save the restart. Restart the server.

12. After the server restarts, rerun the associated probe as shown in step 2c in the Verifying the issue still exists section.

13. If the probe succeeds, failover the databases by running the following command:
    
        Set-MailboxServer server1.contoso.com -DatabaseCopyActivationDisabledAndMoveNow $false

14. If the probe is still failing, you may need assistance to resolve this issue. Contact a Microsoft Support professional to resolve this issue. To contact a Microsoft Support professional, visit the [Exchange Server Solutions Center](http://go.microsoft.com/fwlink/p/?linkid=180809). In the navigation pane, click **Support options and resources** and use one of the options listed under **Get technical support** to contact a Microsoft Support professional. Because your organization may have a specific procedure for directly contacting Microsoft Product Support Services, be sure to review your organization's guidelines first.

</div>

</div>

<div>

## For More Information

[What's new in Exchange 2013](https://technet.microsoft.com/en-us/library/jj150540\(v=exchg.150\))

</div>

</div>

<span> </span>

</div>

</div>

</div>


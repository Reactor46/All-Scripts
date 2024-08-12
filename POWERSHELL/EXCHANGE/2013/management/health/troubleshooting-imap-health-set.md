﻿---
title: Troubleshooting IMAP Health Set
TOCTitle: Troubleshooting IMAP Health Set
ms:assetid: 2a06e671-4cc2-4ce5-bf53-6243ea140f20
ms:mtpsurl: https://technet.microsoft.com/en-us/library/ms.exch.scom.imap(v=EXCHG.150)
ms:contentKeyID: 49720747
ms.date: 10/08/2015
mtps_version: v=EXCHG.150
---

<div data-xmlns="http://www.w3.org/1999/xhtml">

<div class="topic" data-xmlns="http://www.w3.org/1999/xhtml" data-msxsl="urn:schemas-microsoft-com:xslt" data-cs="http://msdn.microsoft.com/en-us/">

<div data-asp="http://msdn2.microsoft.com/asp">

# Troubleshooting IMAP Health Set

</div>

<div id="mainSection">

<div id="mainBody">

<span> </span>

_**Applies to:** Exchange Server 2013_

_**Topic Last Modified:** 2015-03-09_

The IMAP health set monitors the availability of the IMAP4 proxy infrastructure on the Client Access server (CAS). The IMAP health set is closely related to the following health sets:

[Troubleshooting IMAP.Protocol Health Set](troubleshooting-imap-protocol-health-set.md)

[Troubleshooting IMAP.Proxy Health Set](troubleshooting-imap-proxy-health-set.md)

If you receive an alert that specifies that IMAP is unhealthy, this indicates an issue that may prevent your users from accessing their mailboxes by using IMAP.

<span id="EXP"></span>

<div>

## Explanation

The IMAP4 service is monitored by using the following probes and monitors.


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
<td><p>ImapCTPProbe</p></td>
<td><p>IMAP</p></td>
<td><p>Active Directory</p>
<p>Authentication</p>
<p>Mailbox Server Authentication</p>
<p>High Availability</p>
<p>Networking</p></td>
<td><p>ImapCTPMonitor (IMAP health set)</p></td>
</tr>
<tr class="even">
<td><p>ImapProxyTestProbe</p></td>
<td><p>IMAP.Proxy</p></td>
<td><p>Active Directory</p>
<p>Authentication</p></td>
<td><p>ImapProxyTestMonitor(IMAP.Proxy health set)</p></td>
</tr>
<tr class="odd">
<td><p>ImapDeepTestProbe</p></td>
<td><p>IMAP.Protocol</p></td>
<td><p>Active Directory</p>
<p>Authentication</p>
<p>Information Store</p>
<p>High Availability</p></td>
<td><p>IMAP.Protocol (IMAP.Protocol health set)</p></td>
</tr>
<tr class="even">
<td><p>ImapSelfTestProbe</p></td>
<td><p>IMAP.Protocol</p></td>
<td><p>Active Directory</p>
<p>Authentication</p></td>
<td><p>IMAP.Protocol (IMAP.Protocol health set)</p>
<p>AverageCommandProcessingTimeGt60sMonitor (IMAP health set)</p></td>
</tr>
</tbody>
</table>


For more information about probes and monitors, see [Server health and performance](https://technet.microsoft.com/en-us/library/jj150551\(v=exchg.150\)).

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
        
        For example, to retrieve the IMAP health set details about server1.contoso.com, run the following command:
        
            Get-ServerHealth server1.contoso.com | ?{$_.HealthSetName -like "IMAP*"}
    
    2.  Review the command output to determine which monitor reported the error. The **AlertValue** value for the monitor that issued the alert will be `Unhealthy`.
    
    3.  Rerun the associated probe for the monitor that is in an unhealthy state. Refer to the table in the Explanation section to find the associated probe. To do this, run the following command:
        
            Invoke-MonitoringProbe <health set name>\<probe name> -Server <server name> | Format-List
        
        For example, assume that the failing monitor is **ImapCTPMonitor**. The probe associated with that monitor is **ImapCTPProbe**. To run that probe on server1.contoso.com, run the following command:
        
            Invoke-MonitoringProbe IMAP\ImapCTPProbe -Server server1.contoso.com | Format-List
    
    4.  In the command output, review the **Result** value of the probe. If the value is **Succeeded**, the issue was a transient error, and it no longer exists. Otherwise, refer to the recovery steps outlined in the following sections.

</div>

<span id="TestMonitors"></span>

<div>

## ImapTestDeepMonitor and ImapSelfTestMonitor Recovery Actions

1.  Restart the Exchange IMAP4 service on the back-end server. For more information about how to stop and start the IMAP4 service, see [Start and stop the IMAP4 services](https://technet.microsoft.com/en-us/library/bb124022\(v=exchg.150\)).

2.  Rerun the associated probe as shown in step 2c in the Verifying the issue still exists section.

3.  If the issue still exists, you must failover the databases hosted on the mailbox server by using the following command:
    
        Set-MailboxServer -Identity <ServerName> -DatabaseCopyActivationDisabledAndMoveNow $true

4.  Verify that all databases have been moved off the server that’s reporting the issue. To do this, run the following command:
    
        Get-MailboxDatabaseCopyStatus -Server server1.contoso.com | Group Status
    
    If the command output shows no active copies on the server, restart the server.

5.  Rerun the associated probe as shown in step 2c in the Verifying the issue still exists section.

6.  If the probe succeeds, failover the databases by running the following command:
    
        Set-MailboxServer server1.contoso.com -DatabaseCopyActivationDisabledAndMoveNow $false

7.  If the probe continues to fail, you may need assistance to resolve this issue. Contact a Microsoft Support professional to resolve this issue. To contact a Microsoft Support professional, visit the [Exchange Server Solutions Center](http://go.microsoft.com/fwlink/p/?linkid=180809). In the navigation pane, click **Support options and resources** and use one of the options listed under **Get technical support** to contact a Microsoft Support professional. Because your organization may have a specific procedure for directly contacting Microsoft Product Support Services, be sure to review your organization's guidelines first.

</div>

<div>

## ImapCTPMonitor Recovery Actions

This monitor alert is typically issued on CAS servers.

1.  Restart the Exchange IMAP4 service on the backend server. For more information about stopping and starting the IMAP4 service, see [Start and stop the IMAP4 services](https://technet.microsoft.com/en-us/library/bb124022\(v=exchg.150\))

2.  Rerun the associated probe as shown in step 2.c. in the Verifying the issue still exists section.

3.  If the issue still exists, you must turn on IMAP protocol logging to help resolve the issue. To do this, follow these steps:
    
    1.  In Windows PowerShell, run the following command:
        
            Set-ImapSettings -server <CAS server name> -ProtocolLoggingEnabled $true
    
    2.  Restart the Exchange IMAP4 service on the backend server. For more information about how to stop and start the IMAP4 service, see [Start and stop the IMAP4 services](https://technet.microsoft.com/en-us/library/bb124022\(v=exchg.150\))
    
    3.  Rerun the associated probe as shown in step 2c in the Verifying the issue still exists section.
    
    4.  Run the following command, and then determine the location of the log file. To do this, run the following command:
        
            Get-ImapSettings -server <CAS server name>
    
    5.  Determine the mailbox that is serving this command. The name of the mailbox server is the value for `_Mbx:` value in the error message.
    
    6.  Run the following command:
        
            Get-ServerHealth mailbox1.contoso.com | ?{$_.HealtSetName -like "IMAP*"}
        
        **Note**   In this command, replace *mailbox1.contoso.com* with the actual Mailbox server name.
    
    7.  If any of the monitors that are listed in the command output are reported as unhealthy, you must address those monitors first. Follow the troubleshooting steps outlined in the ImapTestDeepMonitor and ImapSelfTestMonitor Recovery Actions section.

4.  If the Mailbox server is reported as healthy, restart the CAS.

5.  After the server restarts, rerun the associated probe as shown in step 2c in the Verifying the issue still exists section.

6.  Turn off protocol logging. To do this, run the following Windows PowerShell command:
    
        Set-ImapSettings -server <CAS server name> -ProtocolLoggingEnabled $false

7.  Restart the IMAP4 service.

8.  If the probe continues to fail, you may need assistance to resolve this issue. Contact a Microsoft Support professional to resolve this issue. To contact a Microsoft Support professional, visit the [Exchange Server Solutions Center](http://go.microsoft.com/fwlink/p/?linkid=180809). In the navigation pane, click **Support options and resources** and use one of the options listed under **Get technical support** to contact a Microsoft Support professional. Because your organization may have a specific procedure for directly contacting Microsoft Product Support Services, be sure to review your organization's guidelines first.

</div>

<div>

## ImapProxyTestMonitor recovery actions

1.  Restart the IMAP4 service.

2.  Rerun the associated probe as shown in step 2c in the Verifying the issue still exists section.

3.  If probe is still failing, restart the CAS.

4.  After the server restarts, rerun the associated probe as shown in step 2c in the Verifying the issue still exists section.

5.  If the probe continues to fail, you may need assistance to resolve this issue. Contact a Microsoft Support professional to resolve this issue. To contact a Microsoft Support professional, visit the [Exchange Server Solutions Center](http://go.microsoft.com/fwlink/p/?linkid=180809). In the navigation pane, click **Support options and resources** and use one of the options listed under **Get technical support** to contact a Microsoft Support professional. Because your organization may have a specific procedure for directly contacting Microsoft Product Support Services, be sure to review your organization's guidelines first.

</div>

<div>

## AverageCommandProcessingTimeGt60sMonitor RequestsQueuedGt500Monitor Recovery Actions

This monitor alert is typically issued on CA and Mailbox servers.

1.  Restart the Exchange IMAP4 service on the back-end server or CAS. For more information about how to stop and start the IMAP4 service, see [Start and stop the IMAP4 services](https://technet.microsoft.com/en-us/library/bb124022\(v=exchg.150\))

2.  Wait 10 minutes to see whether the monitor stays healthy. After 10 minutes, run the following command:
    
        Get-ServerHealth server1.contoso.com | ?{$_.HealthSetName -like "IMAP*"}
    
    **Note**   In this command, replace *server1.contoso.com* with the actual server name.

3.  Wait 10 minutes, and then run the command shown in step 2 again to see whether the monitor stays healthy.

4.  If the issue still exists, you must restart the server. If the server is a CAS, just restart the server. If the server is a Mailbox server, do the following:
    
    1.  Failover the databases that are hosted on the server. To do this, run the following command:
        
            Set-MailboxServer server1.contoso.com -DatabaseCopyActivationDisabledAndMoveNow $true
        
        **Note**   In this and all subsequent code examples, replace *server1.contoso.com* with the actual server name.
    
    2.  Verify that all databases have been moved off the server that is reporting the issue. To do this, run the following command:
        
            Get-MailboxDatabaseCopyStatus -Server server1.contoso.com | Group Status
        
        If the command output shows no active copies on the server, restart the server.

5.  After the server restarts, wait 10 minutes, and then run the command shown in step 2 again to see whether the monitor stays healthy.

6.  If the monitor stays healthy, and if this is a Mailbox server, failover the databases by running the following command:
    
        Set-MailboxServer server1.contoso.com -DatabaseCopyActivationDisabledAndMoveNow $false

7.  If the probe continues to fail, you may need assistance to resolve this issue. Contact a Microsoft Support professional to resolve this issue. To contact a Microsoft Support professional, visit the [Exchange Server Solutions Center](http://go.microsoft.com/fwlink/p/?linkid=180809). In the navigation pane, click **Support options and resources** and use one of the options listed under **Get technical support** to contact a Microsoft Support professional. Because your organization may have a specific procedure for directly contacting Microsoft Product Support Services, be sure to review your organization's guidelines first.

</div>

</div>

<div>

## For More Information

[POP3 and IMAP4](https://technet.microsoft.com/en-us/library/jj657728\(v=exchg.150\))

[Enable IMAP4 in Exchange 2013](https://technet.microsoft.com/en-us/library/bb124489\(v=exchg.150\))

[Test-ImapConnectivity](https://technet.microsoft.com/en-us/library/bb738126\(v=exchg.150\))

</div>

</div>

<span> </span>

</div>

</div>

</div>


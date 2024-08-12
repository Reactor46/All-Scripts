﻿---
title: Troubleshooting EWS Health Set
TOCTitle: Troubleshooting EWS Health Set
ms:assetid: f5aaacdd-7f4a-4d63-8440-1c564e644dfc
ms:mtpsurl: https://technet.microsoft.com/en-us/library/ms.exch.scom.ews(v=EXCHG.150)
ms:contentKeyID: 49720924
ms.date: 10/08/2015
mtps_version: v=EXCHG.150
---

<div data-xmlns="http://www.w3.org/1999/xhtml">

<div class="topic" data-xmlns="http://www.w3.org/1999/xhtml" data-msxsl="urn:schemas-microsoft-com:xslt" data-cs="http://msdn.microsoft.com/en-us/">

<div data-asp="http://msdn2.microsoft.com/asp">

# Troubleshooting EWS Health Set

</div>

<div id="mainSection">

<div id="mainBody">

<span> </span>

_**Applies to:** Exchange Server 2013, Project Server 2013_

_**Topic Last Modified:** 2015-03-09_

The Exchange Web Services (EWS) health set monitors the overall health of the EWS service. The EWS health set is closely related to the following health sets:

[Troubleshooting EWS.Protocol Health Set](troubleshooting-ews-protocol-health-set.md)

[Troubleshooting EWS.Proxy Health Set](troubleshooting-ews-proxy-health-set.md)

If you receive an alert that specifies that EWS is unhealthy, this indicates an issue that may prevent your users from communicating with the Exchange server.

<span id="EXP"></span>

<div>

## Explanation

EWS is monitored by using the following probes and monitors.


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
<td><p>EwsCtpProbe</p></td>
<td><p>EWS</p></td>
<td><p>Information Store</p>
<p>Active Directory Domain Services (AD DS)</p></td>
<td><p>EwsCtpMonitor (EWS health set)</p></td>
</tr>
<tr class="even">
<td><p>EwsSelfTestProbe</p></td>
<td><p>EWS.Protocol</p></td>
<td><p>Active Directory Domain Services (AD DS)</p></td>
<td><p>EWSSelfTestMonitor</p></td>
</tr>
<tr class="odd">
<td><p>EwsDeepTestProbe</p></td>
<td><p>EWS.Protocol</p></td>
<td><p>Information Store</p></td>
<td><p>EWSDeepTestMonitor</p></td>
</tr>
</tbody>
</table>


This probe performs a full EWS logon from the Client Access server (CAS) to a Mailbox server by using a monitoring account. This probe calls the **GetFolder** method on EWS. For more information about probes and monitors, see [Server health and performance](https://technet.microsoft.com/en-us/library/jj150551\(v=exchg.150\)).

</div>

<div>

## Common issues

This probe can fail for any of the following common reasons:

  - A mismatch exists between the authentication mechanism that is used by the probe and the authentication mechanism that is used on the CAS virtual directory.

  - The EWS Application pool in the CAS that’s being monitored is not responding.

  - The CAS is experiencing networking issues when it connects to the Mailbox server.

  - The CAS is experiencing communication issues when it connects to Domain Controllers.

  - The Domain Controllers are not responding.

  - The EWS Application pool that resides on one or more Mailbox servers is not responding.

  - The user’s database is not mounted, or the Information Store is unavailable for a specific mailbox.

  - The Information Store service on one or more Mailbox servers is experiencing issues.

</div>

<div>

## User Action

It's possible that the service recovered after it issued the alert. Therefore, when you receive an alert that specifies that the health set is unhealthy, first verify that the issue still exists. If the issue does exist, perform the appropriate recovery actions outlined in the following sections.

<span id="verify"></span>

<div>

## Verifying the issue still exists

1.  Identify the health set name and the server name in the alert.
    
    When you receive an alert from this HealthSet, the email message contains the following information:
    
    1.  Name of the CAS on which the alert originated
    
    2.  Name of Mailbox server that the probe was monitoring as a target resource
    
    3.  Full exception trace of the last error, including diagnostic data and specific HTTP header information
    
    4.  Time when the incident occurred
    
    5.  Authentication mechanism that was used, and credential information
    
    The exception trace information provides the most important clue as to why the probe is failing. The escalation message also contains the following HTTP headers:
    
    1.  **X-FEServer**   Indicates on which CAS the probe was run
    
    2.  **X-TargetBEServer**   Indicates to which MBX server the request is routed
    
    3.  **X-DiagInfo**   Indicates the MBX server that received the request

2.  The message details provide information about the exact cause of the alert. In most cases, the message details provide sufficient troubleshooting information to identify the root cause. If the message details are not clear, do the following:
    
    1.  Open the Exchange Management Shell, and then run the following command to retrieve the details of the health set that issued the alert:
        
            Get-ServerHealth <server name> | ?{$_.HealthSetName -eq "<health set name>"}
        
        For example, to retrieve the EWS health set details about server1.contoso.com, run the following command:
        
            Get-ServerHealth server1.contoso.com | ?{$_.HealthSetName -eq "EWS"}
    
    2.  Review the command output to determine which monitor reported the error. The **AlertValue** value for the monitor that issued the alert will be `Unhealthy`.
    
    3.  Rerun the associated probe for the monitor that‘s in an unhealthy state. Refer to the table in the Explanation section to find the associated probe. To do this, run the following command:
        
            Invoke-MonitoringProbe <health set name>\<probe name> -Server <server name> | Format-List
        
        For the EWS health set, assume that the failing monitor is **EWSCtpMonitor**. The probe associated with that monitor is **EWSCtpProbe**. To run that probe on server1.contoso.com, run the following command:
        
            Invoke-MonitoringProbe EWS\EWSCtpProbe -Server server1.contoso.com | Format-List
    
    4.  In the command output, review the **Result** value of the probe. If the value is **Succeeded**, the issue was a transient error, and it no longer exists. Otherwise, refer to the recovery steps outlined in the following sections.

</div>

<span id="TestMonitors"></span>

<div>

## EwsCtpMonitor Recovery Actions

1.  Start IIS Manager, and then connect to the server that’s reporting the issue to determine whether the **MSExchangeServicesAppPool** application pool is running on both CA and Mailbox servers.

2.  Locate the MailboxDatabase for the failed probes. The, verify that the Mailbox database is active for the Mailbox server, and that the Information Store is healthy.

3.  Click **Application Pools**, and then recycle the **MSExchangeServicesAppPool** application pool by running the following command from the Shell:
    
        %SystemRoot%\System32\inetsrv\Appcmd recycle MSExchangeServicesAppPool

4.  Rerun the associated probe as shown in step 2c in the Verifying the issue still exists section.

5.  If the issue still exists, recycle the entire IIS service by using the IISReset utility.

6.  Rerun the associated probe as shown in step 2c in the Verifying the issue still exists section.

7.  If the issue still exits, review the protocol log files on the CA and Mailbox servers. The protocol logs for the CAS reside in the *\<exchange server installation directory\>\\Logging\\HttpProxy\\Ews* folder. On the Mailbox server, the logs reside in the *\<exchange server installation directory\>\\Logging\\Ews* folder.

8.  Create a test user account, and then log on by running the test user account against the given CAS. For example, log on by using: https:// \<servername\>/ews/exchange.asmx. If the issue still exists, try a different CAS to determine whether the problem is scoped to that CAS and not to the Mailbox server. If the test user name passes, an issue may affect the specific Mailbox database or Mailbox server on which the monitoring mailbox is located. Try to repeat this step by using a test account that exists in the Mailbox database.

9.  Check network connectivity between the CA and Mailbox server.

10. Check for any alerts on the EWS.Proxy Health Set that might indicate a problem that affects a specific CAS.

11. Check for any alerts on the EWS.Protocol Health Set that might indicate a problem that affects a specific Mailbox server.

12. If the issue still exists, restart the server. To do this, first failover the databases that are hosted on the server. To do this, run the following command:
    
        Set-MailboxServer server1.contoso.com -DatabaseCopyActivationDisabledAndMoveNow $true
    
    **Note**   In this and all subsequent code examples, replace *server1.contoso.com* with the actual server name.

13. Verify that all the databases have been moved off the server that is reporting the issue. To do this, run the following command:
    
        Get-MailboxDatabaseCopyStatus -Server server1.contoso.com | Group Status
    
    If the command output shows no active copies on the server, restart the server.

14. After the server restarts, rerun the associated probe as shown in step 2c in the Verifying the issue still exists section.

15. If the probe succeeds, failover the databases back to the Mailbox server by running the following command:
    
        Set-MailboxServer server1.contoso.com -DatabaseCopyActivationDisabledAndMoveNow $false

16. If the probe is still failing, you may need assistance to resolve this issue. Contact a Microsoft Support professional to resolve this issue. To contact a Microsoft Support professional, visit the [Exchange Server Solutions Center](http://go.microsoft.com/fwlink/p/?linkid=180809). In the navigation pane, click **Support options and resources** and use one of the options listed under **Get technical support** to contact a Microsoft Support professional. Because your organization may have a specific procedure for directly contacting Microsoft Product Support Services, be sure to review your organization's guidelines first.

</div>

</div>

<div>

## For More Information

[Exchange 2013 cmdlets](https://technet.microsoft.com/en-us/library/bb124413\(v=exchg.150\))

[What's new in Exchange 2013](https://technet.microsoft.com/en-us/library/jj150540\(v=exchg.150\))

</div>

</div>

<span> </span>

</div>

</div>

</div>


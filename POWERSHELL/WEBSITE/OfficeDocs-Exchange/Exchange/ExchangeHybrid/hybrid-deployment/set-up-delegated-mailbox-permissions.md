---
title: "Configure Exchange to support delegated mailbox permissions in a hybrid deployment"
ms.author: dstrome
author: dstrome
manager: serdars
ms.date: 1/30/2018
ms.audience: Developer
ms.topic: article
ms.prod: exchange-server-it-pro
localization_priority: Normal
ms.collection: Hybrid
ms.assetid: a2a10cb3-4557-4ff5-8191-c653522f4512
description: "Delegated mailbox permissions enable someone to manage some part of another users's mailbox. A common example of this is an administrative assistant who needs to manage an executive's mailbox and calendar. Hybrid deployments between an on-premises Exchange organization and Office 365 support the Full Access and Send on Behalf of delegated mailbox permissions. However, depending on the version of Exchange you have installed in your on-premises organization, you might need to perform additional configuration to use delegated mailbox permissions in a hybrid deployment. The following lists the versions of Exchange that support delegated mailbox permissions in a hybrid deployment and whether additional configuration is needed for that version."
---

# Configure Exchange to support delegated mailbox permissions in a hybrid deployment

Delegated mailbox permissions enable someone to manage some part of another users's mailbox. A common example of this is an administrative assistant who needs to manage an executive's mailbox and calendar. Hybrid deployments between an on-premises Exchange organization and Office 365 support the **Full Access** and **Send on Behalf of** delegated mailbox permissions. However, depending on the version of Exchange you have installed in your on-premises organization, you might need to perform additional configuration to use delegated mailbox permissions in a hybrid deployment. The following lists the versions of Exchange that support delegated mailbox permissions in a hybrid deployment and whether additional configuration is needed for that version. 
  
- **Exchange 2016** No additional configuration is required. 
    
- **Exchange 2013** A supported Exchange 2013 cumulative update (CU) and additional configuration are required. 
    
- **Exchange 2010** A supported Exchange 2010 update roll (RU) and additional configuration are required. 
    
For more information about the specific requirements to support delegated mailbox permissions in a hybrid deployment, take a look at [Permissions in Exchange hybrid deployments](../permissions.md).
  
The following sections step you through the configuration of Exchange 2013 and Exchange 2010 on-premises deployments to enable support for delegated mailbox permissions. Before you follow these steps, you need to make sure you're on the latest Exchange 2013 CU or Exchange SP3 RU. For more information, see [Hybrid deployment prerequisites](../hybrid-deployment-prerequisites.md).
  
## Exchange 2013

What you need to do to enable support for delegated mailbox permissions depends on a few factors. If you moved mailboxes to Office 365 and at that time:
  
|**The following was installed...**|**And ACLable object synchronization at the organization was...**|**Then you need to...**|
|:-----|:-----|:-----|
|Exchange 2013 CU9 or earlier  <br/> |This feature isn't available in Exchange 2013 CU9 and earlier.  <br/> |Manually configure each mailbox to support ACLs  <br/> |
|Exchange 2013 CU10 or later  <br/> |Disabled  <br/> | Enable ACLable object synchronization at the organization level  <br/>  Manually enable ACLs on each mailbox moved to Office 365 before ACLable object synchronization was enabled at the organization level.  <br/>  No additional configuration is needed for mailboxes moved to Office 365 after ACLable object synchronization is enabled at the organization level.  <br/> |
|Exchange 2013 CU10 or later  <br/> |Enabled  <br/> |No additional configuration is needed  <br/> |
   
### Enable ACLable object synchronization

To enable ACLable object synchronization at the organization level, do the following. 
  
1. Install the latest version of Azure Active Directory Connect (AAD Connect) on all of your AAD Connect servers. This is needed to allow AAD Connect to synchronize the attributes needed to support hybrid permissions. You can download AAD Connect from [Microsoft Azure Active Directory Connect](http://go.microsoft.com/fwlink/p/?LinkID=510956).
    
2. Open the Exchange Management Shell on a server running the latest available Exchange 2013 CU, or the immediately previous CU.
    
3. Run the following command.
    
  ```
  Set-OrganizationConfig -ACLableSyncedObjectEnabled $True
  ```

After you do this, any mailboxes that you move to Office 365 will be properly configured to support delegated mailbox permissions. If mailboxes were moved to Office 365 prior to you completing these steps, you'll need to manually enable ACLs on those mailboxes using the steps in [Enable ACLs on remote mailboxes](set-up-delegated-mailbox-permissions.md#EnableACLs).
  
> [!IMPORTANT]
> ACLs aren't enabled on remote mailboxes created in Office 365. If you create a remote mailbox in Office 365, you'll need to follow the steps in the Enable ACLs on remote mailboxes section to enable ACLs on that remote mailbox. To avoid this extra step, we recommend that you create the mailbox on an on-premises Exchange server and then move the mailbox to Office 365. 
  
### Enable ACLs on remote mailboxes
<a name="EnableACLs"> </a>

To enable ACLs on mailboxes moved to Office 365 before ACLable object synchronization was enabled at the organization level, do the following.
  
1. Open the Exchange Management Shell on a server running the latest available Exchange 2013 CU, or the immediately previous CU.
    
2. To enable ACLs on a single mailbox, run the following command.
    
  ```
  Get-AdUser <Identity> | Set-AdObject -Replace @{msExchRecipientDisplayType=-1073741818}
  ```

3. To enable ACLs on all mailboxes moved to Office 365, run the following command.
    
  ```
  Get-RemoteMailbox | ForEach { Get-AdUser -Identity $_.Guid | Set-ADObject -Replace @{msExchRecipientDisplayType=-1073741818}}
  ```

4. To verify that the mailboxes have been successfully updated, run the following command.
    
  ```
  Get-RemoteMailbox | ForEach { Get-AdUser -Identity $_.Guid -Properties msExchRecipientDisplayType | Format-Table -AutoSize DistinguishedName, msExchRecipientDisplayType}
  ```

## Exchange 2010

Exchange 2010 SP3 servers support the configuration of ACLs on remote mailboxes, however, this configuration needs to be set manually on each mailbox. Unlike newer versions of Exchange, Exchange 2010 doesn't provide the ability to set this feature at the organization level. You need to follow the steps below on any mailboxes that you've previously moved to Office 365, and on any mailboxes that will be moved from an Exchange 2010 SP3 server to Office 365 in the future.
  
### Enable ACLs on remote mailboxes

To enable ACLs on mailboxes moved to Office 365, do the following.
  
1. Open the Exchange Management Shell on a server running the latest available Exchange 2010 SP3 RU, or the immediately previous RU.
    
2. To enable ACLs on a single mailbox, run the following command.
    
  ```
  Get-AdUser <Identity> | Set-AdObject -Replace @{msExchRecipientDisplayType=-1073741818}
  ```

3. To enable ACLs on all mailboxes moved to Office 365, run the following command.
    
  ```
  Get-RemoteMailbox | ForEach { Get-AdUser -Identity $_.Guid | Set-ADObject -Replace @{msExchRecipientDisplayType=-1073741818}}
  ```

4. To verify that the mailboxes have been successfully updated, run the following command.
    
  ```
  Get-RemoteMailbox | ForEach { Get-AdUser -Identity $_.Guid -Properties msExchRecipientDisplayType | Format-Table -AutoSize DistinguishedName, msExchRecipientDisplayType}
  ```



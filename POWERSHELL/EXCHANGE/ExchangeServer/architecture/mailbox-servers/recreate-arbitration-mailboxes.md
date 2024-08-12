---
title: "Recreate missing arbitration mailboxes"
ms.author: dmaguire
author: msdmaguire
manager: serdars
ms.date: 
ms.audience: ITPro
ms.topic: article
ms.prod: exchange-server-it-pro
localization_priority: Normal
ms.assetid: b9004562-b0f2-4460-a623-94883834f73f
description: "Summary: Lean about the built-in arbitration mailboxes in Exchange 2016 and Exchange 2019 and how to recreate them."
---

# Recreate missing arbitration mailboxes

Exchange Server contains five special system mailboxes known as *arbitration mailboxes*. Arbitration mailboxes are used for storing different types of system data and for managing messaging approval workflow. The following table lists each type of arbitration mailbox and their responsibilities.
  
|**Arbitration mailbox Name**|**Display name**|**Persisted capabilities**|**Function**|
|:-----|:-----|:-----|:-----|
|FederatedEmail.4c1f4d8b-8179-4148-93bf-00a95fa1e042|Microsoft Exchange Federation Mailbox|{}|This mailbox stores data used to maintain federation between different Exchange organizations. This includes Rights Management Services, cross-premises mail-flow monitoring probes and responses, notifications, online archives, messaging records management, and cross-premises free/busy information.|
|SystemMailbox{1f05a927-XXXX-XXXX-XXXX-XXXXXXXXXXXX} <br/> (for example, SystemMailbox{1f05a927-9350-4efe-a823-5529c2d64109}; most of the mailbox name is unique to your organization)|Microsoft Exchange Approval Assistant|{}|This mailbox is provisioned for use by the Exchange approval framework for recipient moderation and auto group approval requests.|
|Migration.8f3e7716-2011-43e4-96b1-aba62d229136|Microsoft Exchange Migration|{OrganizationCapabilityManagement}|Stores data for the Exchange migration service to use when moving mailboxes in batches.|
|SystemMailbox{e0dc1c29-89c3-4034-b678-e6c29d823ed9}|Microsoft Exchange|{OrganizationCapabilityUMDataStorage}|Discovery system mailbox. <br/> Provisioned for use by e-Discovery feature, which is used by compliance officers to locate messages that match specified selection criteria. This mailbox is also used by Unified Messaging for storing UM console attending files and other information.|
|SystemMailbox{bb558c35-97f1-4cb9-8ff7-d53741dc928c}|Microsoft Exchange|{OrganizationCapabilityUMGrammarReady, OrganizationCapabilityPstProvider, OrganizationCapabilityMessageTracking, OrganizationCapabilityMailRouting, OrganizationCapabilityClientExtensions, OrganizationCapabilityGMGen, OrganizationCapabilityOABGen, OrganizationCapabilityUMGrammar}|This is known as an organization mailbox. It is used for creating offline address books (OABs). To load-balance OAB generation across your organization, including across geographically separate sites, you can create additional organization mailboxes.|
   
If you need to re-create one of more of these arbitration mailboxes, see the instructions that follow.
  
## What do you need to know before you begin?

- Estimated time to complete: 10 minutes per procedure.
    
- You need to be assigned permissions before you can perform these procedures. To see what permissions you need, see the "Recipient Provisioning Permissions" section in the [Recipients Permissions](../../permissions/feature-permissions/recipient-permissions.md) topic.

- To run Setup.exe /PrepareAD, your account needs to be a member of the Enterprise Admins security group.

- The computer that you use to run Setup.exe /PrepareAD requires access to Setup.exe in the Exchange installation files:

    1. Use your most recently downloaded copy of the Exchange ISO image file, or download an updated copy from [Updates for Exchange Server](../../new-features/updates.md).
    
    2. In File Explorer, right-click on the Exchange ISO image file and then select **Mount**. Note the virtual DVD drive letter that's assigned.

    3. Opwn a Windows Command Prompt window. For example:

      - Press the Windows key + 'R' to open the **Run** dialog, type cmd.exe, and then press **OK**.

      - Press **Start**. In the **Search** box, type **Command Prompt**, then in the list of results, select **Command Prompt**.

- For more information about opening the Exchange Management Shell, see [Open the Exchange Management Shell](https://docs.microsoft.com/powershell/exchange/exchange-server/open-the-exchange-management-shell).

- For more information about running Exchange Setup in unattended mode, see [Use unattended mode in Exchange Setup](../../plan-and-deploy/deploy-new-installations/unattended-installs.md).
    
## Re-create an arbitration mailbox

Use the following instructions to re-create a particular type of arbitration mailbox.
  
### Re-create the Microsoft Exchange Federation Mailbox

To re-create the arbitration mailbox FederatedEmail.4c1f4d8b-8179-4148-93bf-00a95fa1e042, run the following commands:
  
1. If the arbitration mailbox is missing, run the following command from a Windows Command Prompt window:
    
    ```
    <Virtual DVD drive letter>:\Setup.exe /IAcceptExchangeServerLicenseTerms /PrepareAD
    ```

  For example:

    ```
    E:\Setup.exe /IAcceptExchangeServerLicenseTerms /PrepareAD
    ```

2. In the Exchange Management Shell, run the following command:
    
    ```
    Enable-Mailbox -Arbitration -Identity "FederatedEmail.4c1f4d8b-8179-4148-93bf-00a95fa1e042"
    ```

### Re-create the Microsoft Exchange Approval Assistant mailbox

To re-create the arbitration mailbox SystemMailbox{1f05a927-XXXX-XXXX-XXXX-XXXXXXXXXXXX}, run the following commands:
  
1. If the arbitration mailbox is missing, run the following command from a Windows Command Prompt window:
    
    ```
    <Virtual DVD drive letter>:\Setup.exe /IAcceptExchangeServerLicenseTerms /PrepareAD
    ```

    For example:

    ```
    E:\Setup.exe /IAcceptExchangeServerLicenseTerms /PrepareAD
    ```

2. In the Exchange Management Shell, run the following command:
    
    ```
    Get-User | where {$_.Name -like "SystemMailbox{1f05a927*"} | Enable-Mailbox -Arbitration
    ```

### Re-create the Microsoft Exchange Migration mailbox

To re-create the arbitration mailbox Migration.8f3e7716-2011-43e4-96b1-aba62d229136, run the following commands in the Exchange Management Shell:

1. If the arbitration mailbox is missing, run the following command from a Windows Command Prompt window:
    
    ```
    <Virtual DVD drive letter>:\Setup.exe /IAcceptExchangeServerLicenseTerms /PrepareAD
    ```
  
    For example:

    ```
    E:\Setup.exe /IAcceptExchangeServerLicenseTerms /PrepareAD
    ```

2. In the Exchange Management shell, run the following command:
    
    ```
    Enable-Mailbox -Arbitration -Identity "Migration.8f3e7716-2011-43e4-96b1-aba62d229136"
    ```

3. In the Exchange Management Shell, set the Persisted Capabilities (msExchCapabilityIdentifiers) for the mailbox by running the following command:
    
    ```
    Set-Mailbox "Migration.8f3e7716-2011-43e4-96b1-aba62d229136" -Arbitration -Management $true -Force
    ```

### Re-create the Microsoft Exchange Discovery system mailbox

To re-create the arbitration mailbox SystemMailbox{e0dc1c29-89c3-4034-b678-e6c29d823ed9}, run the following command from a Windows Command Prompt window:
  
```
<Virtual DVD drive letter>:\Setup.exe /IAcceptExchangeServerLicenseTerms /PrepareAD
```

For example:

```
E:\Setup.exe /IAcceptExchangeServerLicenseTerms /PrepareAD
```

### Re-create the Microsoft Exchange organization mailbox for OABs

To re-create the arbitration mailbox SystemMailbox{bb558c35-97f1-4cb9-8ff7-d53741dc928c}, run the following commands:

1. If the arbitration mailbox is missing, run the following command from a Windows Command Prompt window:
    
    ```
    <Virtual DVD drive letter>:\Setup.exe /IAcceptExchangeServerLicenseTerms /PrepareAD
    ```

    For example:

    ```
    E:\Setup.exe /IAcceptExchangeServerLicenseTerms /PrepareAD
    ```

2. In the Exchange Management Shell, run the following command:
    
    ```
    Enable-Mailbox -Arbitration -Identity "SystemMailbox{bb558c35-97f1-4cb9-8ff7-d53741dc928c}"
    ```

3. In the Exchange Management Shell, set the Persisted Capabilities (msExchCapabilityIdentifiers) for the mailbox by running the following command:
    
    ```
    Get-Mailbox "SystemMailbox{bb558c35-97f1-4cb9-8ff7-d53741dc928c}" -Arbitration | Set-Mailbox -Arbitration -UMGrammar $true -OABGen $true -GMGen $true -ClientExtensions $true -MessageTracking $true -PstProvider $true -MaxSendSize 1GB -Force
    ```

4. In the Exchange Management Shell, add the required capabilities to the mailbox by running the following commands:

    ```
    $OABMBX = Get-Mailbox "SystemMailbox{bb558c35-97f1-4cb9-8ff7-d53741dc928c}" -Arbitration; Set-ADUser $OABMBX.SamAccountName -Add @{"msExchCapabilityIdentifiers"="40","42","43","44","47","51","52","46"}
    ```

## How do you know this worked?

To verify that you've successfully re-created the arbitration mailbox, set the search scope to search the entire Active Directory forest, an then use the **Get-Mailbox** cmdlet with the _Arbitration_ switch to retrieve system mailboxes.

```
Set-ADServerSettings -ViewEntireForest $true; Get-Mailbox -Arbitration | Format-Table Name,DisplayName
```

View the results of the command to verify that appropriate system mailbox, either by Name or Display Name from the above table, has been re-created.
  
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at: [Exchange Server](https://go.microsoft.com/fwlink/p/?linkId=60612).

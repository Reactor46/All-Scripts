---
title: "Remove an address book policy"
ms.author: kwekua
author: kwekua
manager: scotv
ms.date: 6/24/2018
ms.audience: ITPro
ms.topic: article
ms.service: exchange-online
localization_priority: Normal
ms.assetid: c20c6f82-2f75-4116-9be1-c5af10113f71
description: "Learn how to remove an address book policy (ABP)."
---

# Remove an address book policy

You can only remove an address book policy (ABP) in Exchange Online PowerShell. To connect to Exchange Online PowerShell, see [Connect to Exchange Online PowerShell](https://go.microsoft.com/fwlink/p/?linkid=396554).

For more information about address book policies, see [Address book policies](address-book-policies.md). 
  
## What do you need to know before you begin?

- Estimated time to complete: Less than 5 minutes.
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "Address book policies" entry in the [Email Address and Address Book Permissions](https://technet.microsoft.com/library/1c1de09d-16ef-4424-9bfb-eb7edffbc8c2.aspx) topic. 
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts for the Exchange admin center](../../accessibility/keyboard-shortcuts-in-admin-center.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542) or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351). 
  
## Use Exchange Online PowerShell to remove an ABP

### Step 1: Verify the ABP isn't assigned to a mailbox

You can't remove an ABP if it's assigned to a mailbox.

To see if the ABP is assigned to an active mailbox, replace \<AddressBookPolicyName\> with the name of the ABP and run the following command:

```
Get-Mailbox | Where $._AddressBookPolicy -eq "<AddressBookPolicyName>"
```

To see if the ABP is assigned to a soft-deleted mailbox, replace \<AddressBookPolicyName\> with the name of the ABP and run the following command:

```
Get-Mailbox -SoftDeletedMailbox | Where $._AddressBookPolicy -eq "<AddressBookPolicyName>"
```

To remove the ABP assignment from mailboxes, replace \<AddressBookPolicyName\> with the name of the ABP and run one or both of the following commands:

```
Get-Mailbox | Where $._AddressBookPolicy -eq "<AddressBookPolicyName>" | Set-Mailbox -AddressBookPolicy $null
```

```
Get-Mailbox -SoftDeletedMailbox | Where $._AddressBookPolicy -eq "<AddressBookPolicyName>" | Set-Mailbox -AddressBookPolicy $null
```

### Step 2: Remove the ABP

This example removes the ABP named ABP TailspinToys.
  
```
Remove-AddressBookPolicy -Identity "ABP TailspinToys"
```

For detailed syntax and parameter information, see [Remove-AddressBookPolicy](https://technet.microsoft.com/library/57ff215a-cba5-46d1-a7f7-ab2512ce4b6f.aspx).
  


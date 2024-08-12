---
title: "Manage user access to add-ins for Outlook"
ms.author: dmaguire
author: msdmaguire
manager: laurawi
ms.date: 7/11/2018
ms.audience: ITPro
ms.topic: article
ms.service: exchange-online
localization_priority: Normal
ms.assetid: e5833dec-a23a-439e-ac03-92671817bff8
description: "You can use the Exchange admin center (EAC) or Exchange Online PowerShell to manage user access to add-ins for Outlook."
---

# Manage user access to add-ins for Outlook

You can use the Exchange admin center (EAC) or Exchange Online PowerShell to manage user access to add-ins for Outlook.
  
- Using the EAC, you can manage basic add-in access settings for your users at an organizational level. For example, you can configure whether an add-in is enabled or disabled for your users. You can also specify whether an add-in is required or optional for your users.
    
- With Exchange Online PowerShell, you can manage all the settings that you can with the EAC, as well as other settings. For example, you can limit availability to specific users in your organization.
    
For additional management tasks, see [Add-ins for Outlook](add-ins-for-outlook.md).
  
## What do you need to know before you begin?

- Estimated time to complete: 5 minutes.
    
- For more information about the EAC, see [Exchange admin center in Exchange Online](../../exchange-admin-center.md).
    
- To learn how to connect to Exchange Online PowerShell, see [Connect to Exchange Online PowerShell](https://go.microsoft.com/fwlink/p/?linkid=396554).
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "Add-ins for Outlook" entry in the [Recipients permissions](https://technet.microsoft.com/library/5b690bcb-c6df-4511-90e1-08ca91f43b37.aspx) topic. 
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts for the Exchange admin center](../../accessibility/keyboard-shortcuts-in-admin-center.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542) or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351).. 
  
## Specify whether an add-in is available, enabled, or disabled

### Use the EAC to specify whether an add-in is available, enabled, or disabled

1. In the EAC, navigate to **Organization** \> **Add-ins**.
    
2. In the list view, select the add-in that you want to change settings for, and then click **Edit** ![Edit icon](../../media/ITPro_EAC_EditIcon.gif).
    
3. If you don't want your users to use the add-in, clear the **Make this add-in available to users in your organization** check box, and then click **Save**.
    
4. If you want your users to be able to use the add-in, select **Make this add-in available to users in your organization**, and then select the option you want.
    
  - **Optional, enabled by default**: Use this setting if you want to allow your users to turn off the add-in.
    
  - **Optional, disabled by default**: Use this setting if you want to allow your users to turn on the add-in.
    
  - **Mandatory, always enabled. Users can't disable this add-in**: Use this setting if you don't want your users to turn off the add-in.
    
5. Click **Save**.
    
### Use Exchange Online PowerShell to specify whether an add-in is available, enabled, or disabled

First, run the following command to find the display names and add-in IDs for all the add-ins for Outlook installed for your organization.
  
```
Get-App -OrganizationApp | Format-List DisplayName,AppId
```

The **AppId** value is a GUID that uniquely identifies the add-in (for example, `fe93bfe1-7947-460a-a5e0-7a5906b51360`). You use the **AppId** value to identify and change the settings of the add-in. 
  
To disable and hide an add-in from all your users, replace _\<AppId\>_ with the real **AppId** value and run the following command: 
  
```
Set-App -Identity <AppId> -OrganizationApp -Enabled $false
```

To enable an add-in by default, but allow your users to turn it off, replace _\<AppId\>_ with the real **AppId** value and run the following command: 
  
```
Set-App -Identity <AppId> -OrganizationApp -Enabled $true -DefaultStateForUser Enabled
```

To disable an add-in by default, but allow your users to turn it on, replace _\<AppId\>_ with the real **AppId** value and run the following command: 
  
```
Set-App -Identity <AppId> -OrganizationApp -Enabled $true -DefaultStateForUser Disabled
```

If you want the add-in to be required for your users, replace _\<AppId\>_ with the real **AppId** value and run the following command: 
  
```
Set-App -Identity <AppId> -OrganizationApp -Enabled $true -DefaultStateForUser AlwaysEnabled
```

For detailed syntax and parameters, see [Set-app](https://technet.microsoft.com/library/3506b2b9-dc23-4ed9-84f5-8839c4c3c974.aspx).
  
### How do you know this worked?

To verify that you've successfully configured an add-in, use either of the following steps:
  
- In the EAC, navigate to **Organization** \> **Add-ins** and review the values in the **User Default** and **Provided To** columns. 
    
- In Exchange Online PowerShell, run the following command and verify the values of the **DefaultStateForUser** and **Enabled** properties: 
    
  ```
  Get-App -OrganizationApp | Format-List DisplayName,AppId,Enabled,DefaultStateForUser
  ```

## Use Exchange Online PowerShell to limit add-in availability to specific users

To limit the availability of an add-in to specific users, you can't use the EAC. You can only use Exchange Online PowerShell.
  
This example limits the LinkedIn add-in with the hypothetical **AppId** value `ac83a9d5-5af2-446f-956a-c583adc94d5e` to the members of the distribution group named Marketing. 
  
```
$a = Get-DistributionGroupMember Marketing
```

```
Set-App -Identity ac83a9d5-5af2-446f-956a-c583adc94d5e -OrganizationApp -ProvidedTo SpecificUsers -UserList $a.Identity -DefaultStateForUser Enabled
```

For detailed syntax and parameters, see [Set-app](https://technet.microsoft.com/library/3506b2b9-dc23-4ed9-84f5-8839c4c3c974.aspx).
  
### How do you know this worked?

To verify that you've successfully limited add-in availability to specific users, run the following command in Exchange Online PowerShell and verify the value of the **ProvidedTo** and **UserList** properties: 
  
```
Get-App -OrganizationApp | Format-List DisplayName,AppId,Enabled,DefaultStateForUser,ProvidedTo,UserList
```



---
title: "Delete a UM IP gateway"
ms.author: tonysmit
author: tonysmit
manager: scotv
ms.date: 11/17/2014
ms.audience: ITPro
ms.topic: article
ms.service: exchange-online
localization_priority: Normal
ms.assetid: 569d3741-67dd-4597-8d28-010011be0c12
description: "When you delete a Unified Messaging (UM) IP gateway, Exchange servers can no longer accept incoming calls from the Voice over IP (VoIP) gateway, Session Initiation Protocol (SIP)-enabled Private Branch eXchange (PBX), IP PBX, or session border controller (SBC) associated with the UM IP gateway."
---

# Delete a UM IP gateway

When you delete a Unified Messaging (UM) IP gateway, Exchange servers can no longer accept incoming calls from the Voice over IP (VoIP) gateway, Session Initiation Protocol (SIP)-enabled Private Branch eXchange (PBX), IP PBX, or session border controller (SBC) associated with the UM IP gateway.
  
> [!IMPORTANT]
> You should delete a UM IP gateway only when you fully understand the implications of disabling communication with a VoIP gateway, IP PBX, or SBC. 
  
For additional tasks related to UM IP gateways, see [UM IP gateway procedures](um-ip-gateway-procedures.md).
  
## What do you need to know before you begin?

- Estimated time to complete: Less than 1 minute.
    
- You need to be assigned permissions before you can perform this procedure or procedures. To see what permissions you need, see the "UM IP gateways" entry in the [Unified Messaging Permissions](https://technet.microsoft.com/library/d326c3bc-8f33-434a-bf02-a83cc26a5498.aspx) topic. 
    
- Before you perform these procedures, confirm that a UM dial plan has been created. For detailed steps, see [Create a UM dial plan](create-um-dial-plan.md).
    
- Before you perform these procedures, confirm that a UM IP gateway has been created. For detailed steps, see [Create a UM IP gateway](create-um-ip-gateway.md).
    
- For information about keyboard shortcuts that may apply to the procedures in this topic, see [Keyboard shortcuts for the Exchange admin center](../../accessibility/keyboard-shortcuts-in-admin-center.md).
    
> [!TIP]
> Having problems? Ask for help in the Exchange forums. Visit the forums at [Exchange Online](https://go.microsoft.com/fwlink/p/?linkId=267542) or [Exchange Online Protection](https://go.microsoft.com/fwlink/p/?linkId=285351).. 
  
## Use the EAC to delete a UM IP gateway

1. In the EAC, navigate to **Unified Messaging** \> **UM IP Gateways**, select the UM IP gateway you want to delete, and then click **Delete** ![Delete icon](../../media/ITPro_EAC_DeleteIcon.gif).
    
2. On the **Warning** page, click **Yes**.
    
## Use Exchange Online PowerShell to delete a UM IP gateway

This example deletes the UM IP gateway named `MyUMIPGateway`.
  
```
Remove-UMIPGateway -Identity MyUMIPGateway
```



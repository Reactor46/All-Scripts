---
title: "Common attachment blocking scenarios for mail flow rules"
ms.author: chrisda
author: chrisda
manager: laurawi
ms.date: 7/11/2018
ms.audience: ITPro
ms.topic: article
ms.service: exchange-online
localization_priority: Normal
ms.assetid: 5c576439-d55b-4c7f-90ed-a7f72cbb16c2
description: "Your organization might require that certain types of messages be blocked or rejected in order to meet legal or compliance requirements, or to implement specific business needs. This article discusses examples of common scenarios for blocking all attachments which you can set up using transport rules in Exchange."
---

# Common attachment blocking scenarios for mail flow rules

Your organization might require that certain types of messages be blocked or rejected in order to meet legal or compliance requirements, or to implement specific business needs. This article discusses examples of common scenarios for blocking all attachments which you can set up using transport rules in Exchange.
  
For additional examples showing how to block specific attachments, see:
  
- [Using transport rules to inspect message attachments](https://technet.microsoft.com/library/c0de687e-e33c-4e8a-b253-771494678795.aspx) (Exchange Server) 
    
- [Use mail flow rules to inspect message attachments in Office 365](inspect-message-attachments.md) (Exchange Online, Exchange Online Protection) 
    
The malware filter includes a Common Attachment Types Filter. In the Exchange admin center (EAC), go to **Protection**, then click **New** ( ![Add Icon](../../media/ITPro_EAC_AddIcon.gif)) to add filters. In the Exchange Online portal, browse to **Protection**, and then select **Malware Filter**.
  
To get started implementing any of these scenarios to block certain message types:
  
1. Sign in to the Exchange admin center.
    
2. Go to **Mail flow** \> **Rules**.
    
3. Click **New** ( ![Add Icon](../../media/ITPro_EAC_AddIcon.gif)) and then select **Create a new rule**.
    
4. In the **Name** box, specify a name for the rule, and then click **More options**.
    
5. Select the conditions and actions you want.
    
 **Note**: In the EAC, the smallest attachment size that you can enter is 1 kilobyte, which should detect most attachments. However, if you want to detect every possible attachment of any size, you need to use PowerShell to adjust the attachment size to 1 byte after you create the rule in the EAC. To learn how to connect to Exchange Online PowerShell, see [Connect to Exchange Online PowerShell](https://go.microsoft.com/fwlink/p/?linkid=396554). To learn how to connect to Exchange Online Protection PowerShell, see [Connect to Exchange Online Protection PowerShell](https://go.microsoft.com/fwlink/p/?linkid=627290).
  
Replace _\<Rule Name\>_ with the name of the existing rule, and run the following command to set the attachment size to 1 byte: 
  
```
Set-TransportRule -Identity "<Rule Name>" -AttachmentSizeOver 1B
```

After you adjust the attachment size to 1 byte, the value that's displayed for the rule in the EAC is **0.00 KB**.
  
## Example 1: Block messages with attachments, and notify the sender
<a name="ex1"> </a>

If you don't want people in your organization to send or receive attachments, you can set up a transport rule to block all messages with attachments.
  
In this example, all messages sent to or from the organization with attachments are blocked.
  
![Rule that blocks all attachments](../../media/38094183-166f-4ba5-a9cf-242e7d0f4e04.png)
  
If all you want to do is block the message, you might want to stop rule processing once this rule is matched. Scroll down the rule dialog box, and select the **Stop processing more rules** check box. 
  
## Example 2: Notify intended recipients when an inbound message is blocked
<a name="ex2"> </a>

If you want to reject a message but let the intended recipient know what happened, you can use the **Notify the recipient with a message** action. 
  
You can include placeholders in the notification message so that it includes information about the original message. The placeholders must be enclosed in two percent signs (%%), and when the notification message is sent, the placeholders are replaced with information from the original message. You can also use basic HTML such as \<br\>, \<b\>, \<i\>, and \<img\> in the message.
  

|**Type of information**|**Placeholder**|
|:-----|:-----|
|Sender of the message.|%%From%% |
|Recipients listed on the "To" line.|%%To%% |
|Recipients listed on the "Cc" line.|%%Cc%% |
|Subject of the original message.|%%Subject%% |
|Headers from the original message. This is similar to the list of headers in a delivery status notification (DSN) generated for the original message.|%%Headers%% |
|Date the original message was sent.|%%MessageDate%% |
   
In this example, all messages that contain attachments and are sent to people inside your organization are blocked, and the recipient is notified.
  
![Rule that notifies recipients when an inbound message is blocked](../../media/f9a14733-d68a-4528-a736-206325881c47.png)
  
## Example 3: Modify the subject line for notifications
<a name="ex3"> </a>

When a notification is sent to the recipient, the subject line is the subject of the original message. If you want to modify the subject so that it is clearer to the recipient, you must use two transport rules:
  
- The first rule adds the word "undeliverable" to the beginning of the subject of any messages with attachments.
    
- The second rule blocks the message and sends a notification message to the sender using the new subject of the original message.
    
> [!IMPORTANT]
> The two rules must have identical conditions. Rules are processed in order, so the first rule adds the word "undeliverable", and the second rule blocks the message and notifies the recipient. 
  
Here's what the first rule would look like if you want to add "undeliverable" to the subject:
  
![Rule that prepends Undeliverable to messages with attachments](../../media/2552b0bd-c69d-48b4-9e69-267fcaf20e70.png)
  
And the second rule does the blocking and notification (the same rule from Example 2):
  
![Rule that notifies recipients when an inbound message is blocked](../../media/f9a14733-d68a-4528-a736-206325881c47.png)
  
## Example 4: Apply a rule with a time limit
<a name="ex4"> </a>

If you have a malware outbreak, you might want to apply a rule with a time limit so that you temporarily block attachments. For example, the following rule has both a start and stop day and time:
  
![Rule showing a time limit](../../media/bdc8c4d8-72fa-4c5b-97f2-5fe76d50e643.png)
  
## See also
<a name="ex4"> </a>

[Mail flow rules (transport rules) in Exchange Online](mail-flow-rules.md)

[Transport Rules (Exchange Server 2016)](https://technet.microsoft.com/library/c3d2031c-fb7b-4866-8ae1-32928d0138ef.aspx)
  
[Transport rules (Exchange Online Protection)](https://technet.microsoft.com/library/9c2cf227-eff7-48ef-87fb-487186e47363.aspx)


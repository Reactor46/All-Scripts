---
title: "Configure mail flow using connectors in Office 365"
ms.author: supotter
author: supotter
manager: serdars
ms.date: 6/23/2018
ms.audience: ITPro
ms.topic: article
f1_keywords:
- 'ms.exch.eac.ConnectorSelection'
ms.service: exchange-online
localization_priority: Normal
ms.assetid: 854b5a50-4462-4836-a092-37e208d29624

description: "Control the flow of email to and from your Office 365 organization. Learn how to use connectors with either Microsoft Exchange Online or Exchange Online Protection."
---

# Configure mail flow using connectors in Office 365

Learn how to use connectors with either Microsoft Exchange Online or Exchange Online Protection. Connectors are a collection of instructions that customize the way your email flows to and from your Office 365 organization.  Actually, most organizations that use Office 365 don't need connectors and this topic will help you decide which one, if needed.
  
## What do connectors do?
<a name="connectorsdo1"> </a>

Set up connectors to:
  
- Enable mail flow between Office 365 and your organization's email servers (also known as an on-premises servers).
    
- Apply security restrictions, or controls, to mail exchanges with a business partner or service provider.
    
- Enable email notifications from a printer or other non-mailbox entity.

- Avoid graylisting/IP throttling
    
> [!NOTE]
> Graylisting is a practice used to protect email systems from spam. In Office 365, graylisting is done by throttling IPs to limit senders from sending suspiciously large amounts of email. Office 365 responds to these abnormal influxes of mail by returning a temporary NDR error, resembling this range: 451 4.7.500-699 (ASxxx). For more details on these types of delivery issues, see [Fix email delivery issues for error code 451 4.7.500-699 (ASxxx) in Office 365](https://support.office.com/article/51356082-9fef-4639-a18a-fc7c5beae0c8). 
  
## What happened to inbound and outbound connectors?
<a name="InboundOutbound1"> </a>

If you previously set up inbound and outbound connectors, they will still function in exactly the same way. The process for setting up connectors has changed; instead of using the terms "inbound" and "outbound", we ask you to specify the start and end points you want to use for a mail flow connector. The way connectors work in the background is the same as before.
  
## When do I need a connector?
<a name="NeedConnector1"></a>

Microsoft Exchange Online is ready to send and receive email from the Internet right away. You don't need to set up connectors unless you have EOP or other specific circumstances, which are described in the table below. Use this table to decide whether you need to set up connectors.
  

|**Scenario**|**What does this mean?**|**Connector required?**|**When creating the connectors, select these options**|
|:-----|:-----|:-----|:-----|
|You have a standalone Exchange Online Protection (EOP) subscription.|You have your own email servers (also called on-premises servers), and you subscribe to EOP only for email protection services. <br/> For details, check **Exchange Online Protection overview** and [How do Office 365 connectors work with my own email servers (also called "on-premises servers")?](set-up-connectors-to-route-mail.md#HowdoconnectorsinEOP).|Yes|**Connector for incoming email:** <br/>• From: Your organization's email server <br/>• To: Office 365  <br/> **Connector for outgoing email**: <br/>• From: Office 365 <br/>• To: Your organization mail server|
|You have an Exchange Online subscription, and some of your mailboxes are on your email servers.|Some of your mailboxes are in Microsoft Exchange Online, and some are on your email servers (also called on-premises servers). Before you set up connectors, check whether you only need connectors or if an Exchange hybrid deployment better meets your business needs. <br/> For details, check [What if I have EOP or Exchange Online and my own email servers?](use-connectors-to-configure-mail-flow.md#WhatifIhave1) and [Exchange Server Hybrid Deployments](https://technet.microsoft.com/library/59e32000-4fcf-417f-a491-f1d8f9aeef9b.aspx).|Yes|**Connector for incoming email:** <br/>• From: Your organization's email server <br/>• To: Office 365  <br/> **Connector for outgoing email:** <br/>• From: Office 365 <br/>• To: Your organization's email server|
|You have all your mailboxes in Exchange Online and you want to send email from other on-premise sources.|You don't have email servers (also called on-premises servers) but you want to let people send email messages from non-mailboxes: printers, fax machines, apps or other devices.  <br/> For details, check [How to set up a multifunction device or application to send email using Office 365](https://technet.microsoft.com/library/2c0012a9-7e71-40cd-a516-1b26117bf491.aspx).|Optional|**Only one connector needed:** <br/>• From: Your organization's email server <br/>• To: Office 365|
|You often exchange email with business partners, and you want to apply certain security restrictions.|When your users exchange email messages with people in partner organizations, you want to make sure that any shared sensitive information is protected. You can do this by using Transport Layer Security (TLS) or by limiting the mail's source destination. <br/> For details, check [Set up connectors for secure mail flow with a partner organization](set-up-connectors-for-secure-mail-flow-with-a-partner.md).|Optional|**Connector for incoming email:** <br/>• From: Partner organization <br/>• To: Office 365  <br/> **Connector for outgoing email:** <br/>• From: Office 365 <br/>• To: Partner organization|
   
> [!TIP]
> If you do not have Microsoft Exchange Online or EOP and are looking for Exchange connectors that apply to Exchange Server 2016 (on-premises server), see [Connectors](https://technet.microsoft.com/library/73559b0c-fc0e-41fd-84df-d07442137a0c.aspx) for information. 
  
## What if I have EOP or Exchange Online and my own email servers?
<a name="WhatifIhave1"> </a>

If you have EOP or Exchange Online and your own email servers (also called on-premises servers), you definitely need connectors. This is more complicated and has more options; here's a breakdown:
  
|**You're running**|**Your service subscription is**|**Have you completed an Exchange hybrid deployment?**|**Do I need to set up connectors manually?**|
|:-----|:-----|:-----|:-----|
|Exchange 2010 or later|Exchange Online Protection|Not available|Yes. Follow the instructions in [Set up connectors to route mail between Office 365 and your own email servers](set-up-connectors-to-route-mail.md).|
|Exchange 2010 or later|Office 365 with Exchange Online|No|Consider whether an Exchange hybrid deployment will better meet your organization's needs by reviewing the topic that matches your current situation in [Exchange Server Hybrid Deployments](https://docs.microsoft.com/exchange/exchange-hybrid). <br/> If a hybrid deployment is the right option for your organization, use the Exchange [Hybrid Configuration wizard](https://docs.microsoft.com/exchange/hybrid-configuration-wizard) to integrate Exchange Online with your on-premises Exchange Server. <br/> If you only want connectors that enable mail routing, follow the instructions in [Set up connectors to route mail between Office 365 and your own email servers](set-up-connectors-to-route-mail.md).|
|Exchange 2010 or later|Office 365 with Exchange Online|Yes|No. The Exchange Hybrid Configuration wizard creates connectors for you. To view or edit those connectors, go to the **Connectors** page in the Exchange admin center (EAC), or rerun the Hybrid Configuration wizard.|
|Exchange Server 2007 or earlier|Exchange Online Protection  <br/> Office 365 with Exchange Online|Not available|Yes. Set them up by following the instructions in [Set up connectors to route mail between Office 365 and your own email servers](set-up-connectors-to-route-mail.md). <br/> In limited circumstances, you might have a hybrid configuration with Microsoft Exchange Server 2007 and Office 365. Check whether connectors are already set up for your organization. To check, go to the **Connectors** page in the EAC.|
|Non-Microsoft SMTP server|Exchange Online Protection <br/> Office 365 with Exchange Online|Not available|Yes. Set them up by following the instructions in [Set up connectors to route mail between Office 365 and your own email servers](set-up-connectors-to-route-mail.md).|
   
### How do connectors in EOP or Exchange Online work with my own email servers (also called on-premises servers)?
<a name="HowDoConnectors1"> </a>

If you have EOP and your own email servers, or if some of your mailboxes are in Exchange Online and some are on your email servers, connectors enable mail flow in both directions. You can enable mail flow between Office 365 and any SMTP-based email server such as Microsoft Exchange, or a third-party email server. Create connectors to enable mail flow in both directions.
  
The diagram below shows how connectors in Office 365 (including Exchange Online or EOP) work with your own email servers.
  
![Connectors between Office 365 and your e-mail server](../../media/0df5ec3d-29c1-4add-9e22-5b0c26bec750.png)
  
In this example, John and Bob are both employees at your company. John has a mailbox on an email server that you manage, and Bob has a mailbox in Office 365. John and Bob both exchange mail with Sun, a customer with an Internet mail account:
  
- When email is sent between John and Bob, connectors are needed
    
- When email is sent between John and Sun, connectors are needed. (All Internet email is delivered via Office 365.)
    
- When email is sent between Bob and Sun, no connector is needed.
    
### What if I have already run the Exchange Hybrid Configuration Wizard?
<a name="HowDoConnectors1"> </a>

If you have already run the Hybrid Configuration wizard, the connectors that you need are already set up for you. You can view your hybrid connectors on the **Connectors** page in the EAC. You can view, troubleshoot, and update these connectors using the procedures described in [Set up connectors to route mail between Office 365 and your own email servers](set-up-connectors-to-route-mail.md), or you can re-run the Hybrid Configuration wizard to make changes.
  
## Connectors for mail flow with a partner organization
<a name="ConnMailFlow1"> </a>

You can create connectors to add additional security restrictions for email sent between Office 365 and a partner organization. A partner can be an organization you do business with, such as a bank. It can also be a cloud email service provider that provides services such as archiving, anti-spam, and so on. You can create a partner connector that defines boundaries and restrictions for email sent to or received from your partners, including scoping the connector to receive email from specific IP addresses, or requiring Transport Layer Security (TLS) encryption.
  
### Example use of connectors with a partner organization

The diagram below shows an example where ContosoBank.com is a business partner that you share financial details with via email. Because you are sharing financial information, you want to protect the integrity of the mail flow between your businesses. Connectors with TLS encryption enable a secure and trusted channel for communicating with ContosoBank.com. In this example, two connectors are created in Office 365. TLS is required for mail flow in both directions, so ContosoBank.com must have a valid encryption certificate. A certificate signed by a certification authority (CA) is recommended.
  
![Connectors between Office 365 and a partner organization](../../media/0f9319ae-84bb-4b05-a79f-12fb988f1d10.png)
  
### Additional partner organization connector options: specify a domain or IP address ranges

When you create a connector, you can also specify the domain or IP address ranges that your partner sends mail from. If email messages don't meet the security conditions that you set, the connector will reject them. For more information about creating connectors to exchange secure email with a partner organization, see [Set up connectors for secure mail flow with a partner organization](set-up-connectors-for-secure-mail-flow-with-a-partner.md).
  
## Connectors for mail notifications from a device including printers
<a name="ConnectorsForMailNot1"> </a>

This scenario applies only to organizations that have all their mailboxes in Microsoft Exchange Online and allows a program or a device, such as a printer, to send email. For example, if you want a printer to send notifications when a print job is ready, or you want your scanner to email documents, you can use this option to send mail through Office 365. For details, see [How to Allow a Multi-function Device or Application to Send E-mail through Office 365 Using SMTP](https://technet.microsoft.com/library/2c0012a9-7e71-40cd-a516-1b26117bf491.aspx).
  
## How do I set up connectors?
<a name="HowDoI1"> </a>

Before you set up a connector, you must set up the accepted domains that you want to define for Office 365. See [Manage accepted domains in Exchange Online](../../mail-flow-best-practices/manage-accepted-domains/manage-accepted-domains.md) for more details. 
  
Connector setup topics:
  
- [Set up connectors to route mail between Office 365 and your own email servers](set-up-connectors-to-route-mail.md)
    
- [Set up connectors for secure mail flow with a partner organization](set-up-connectors-for-secure-mail-flow-with-a-partner.md)
    
## Tell us what you think
<a name="HowDoI1"> </a>

Want to help us improve connectors in Office 365? [Send us feedback](https://go.microsoft.com/fwlink/p/?LinkId=525915), and let us know what you liked, didn't like, or what we can do to make your experience better.
  
## See also
<a name="HowDoI1"> </a>

[Set up connectors to route mail between Office 365 and your own email servers](set-up-connectors-to-route-mail.md)
  
[Mail flow best practices for Exchange Online and Office 365 (overview)](../mail-flow-best-practices.md)

[Set up connectors for secure mail flow with a partner organization](set-up-connectors-for-secure-mail-flow-with-a-partner.md)
  
[What happens when I have multiple connectors for the same scenario?](set-up-connectors-to-route-mail.md#multipleconnectors)


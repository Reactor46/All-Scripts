---
title: "Availability service in Exchange Server"
ms.author: dmaguire
author: msdmaguire
manager: serdars
ms.date: 4/19/2018
ms.audience: ITPro
ms.topic: article
ms.prod: exchange-server-it-pro
localization_priority: Normal
ms.assetid: 9722dea2-2bf8-437c-85c0-3ab65b8a07b9
description: "Summary: Learn about the Availability service in Exchange 2016 and Exchange 2019."
---

# Availability service in Exchange Server

The Availability service makes free/busy information available to Outlook and Outlook on the web (formerly known as Outlook Web App) clients. The Availability service improves information workers' calendaring and meeting scheduling experience by providing secure, consistent, and up-to-date free/busy information.
  
Outlook and Outlook on the web use the Availability service to perform the following tasks:
  
- Retrieve current free/busy information for Exchange mailboxes
    
- Retrieve current free/busy information from other Exchange organizations
    
- Retrieve published free/busy information from public folders for mailboxes on previous versions of Exchange
    
- View attendee working hours
    
- Show meeting time suggestions
    
## How the availability service works in Exchange Server
The Availability service retrieves free/busy information directly from the target Exchange mailbox.
  
Outlook uses the Exchange Autodiscover service to obtain the URL of the Availability service. For more information about the Autodiscover service, see [Autodiscover service](autodiscover.md).
  
You can use the Exchange Management Shell to configure the Availability service. You can't use the Exchange admin center (EAC) to configure the Availability service.
  
The Availability service API is available as a web service to let developers write third-party integration tools.
  
## Availability service and automatic reply messages
The Availability service provides access to automatic-reply messages that users send when they are out of the office or away for an extended period of time.
  
Information workers use the Automatic Replies feature (formerly known as Out of Office) in Outlook and Outlook on the web to alert others when they're unavailable to respond to email messages. This functionality makes it easier to set and manage automatic reply messages for both information workers and administrators.
  
## Methods used to retrieve free/busy information
The following table lists the methods used to retrieve free/busy information in different single-forest topologies.
  
|**Client**|**Source mailbox retrieving free/busy information**|**Target mailbox**|**Free/busy retrieval method**|
|:-----|:-----|:-----|:-----|
|Outlook 2010 or later  <br/> |Exchange 2010 or later <br/> |Exchange 2010 or later <br/> |The Availability service reads free/busy information from the target mailbox.  <br/> |
|Outlook on the web or Outlook Web App <br/> |Exchange 2010 or later <br/> |Exchange 2010 or later <br/> |Outlook on the web or Outlook Web App calls the Availability service API, which reads the free/busy information from the target mailbox.  <br/> |
   


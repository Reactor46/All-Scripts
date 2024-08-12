---
title: "Default settings for Exchange virtual directories"
ms.author: dmaguire
author: msdmaguire
manager: serdars
ms.date: 7/5/2018
ms.audience: ITPro
ms.topic: article
ms.prod: exchange-server-it-pro
localization_priority: Normal
ms.assetid: d2d89ce6-4721-4737-a325-fba5ad9422e0
description: "Summary: Learn about the default Client Access virtual directory settings on Mailbox servers in Exchange 2016 and Exchange 2019."
---

# Default settings for Exchange virtual directories

Exchange Server 2016 and Exchange Server 2019 automatically configure multiple Internet Information Services (IIS) virtual directories during the server installation. The tables in the following sections show the settings for the Client Access (frontewnd) services on Mailbox servers and the default IIS authentication and Secure Sockets Layer (SSL) settings.

## Client Access services on Mailbox servers

The following table lists the default settings on an Exchange Mailbox server that's running Client Access services.

**Default Mailbox server running Client Access services IIS authentication and SSL settings**

|**Virtual directory**|**Authentication method**|**SSL settings**|**Management method**|
|:-----|:-----|:-----|:-----|
|Default website|Anonymous|Required|IIS management console|
|aspnet_client|Anonymous authentication|SSL required <br/> Requires 128-bit encryption|IIS management console|
|Autodiscover|Anonymous authentication <br/> Basic authentication <br/> Windows authentication|SSL requiredRequires 128-bit encryption|Exchange Management Shell|
|ecp|Anonymous authentication <br/> Basic authentication|SSL required <br/> Requires 128-bit encryption|Exchange admin center (EAC) or Exchange Management Shell|
|EWS|Anonymous authentication <br/> Windows authentication|SSL required <br/> Requires 128-bit encryption|Exchange Management Shell|
|Microsoft-Server-ActiveSync|Basic authentication|SSL required <br/> Requires 128-bit encryption|EAC or Exchange Management Shell|
|OAB|Windows authentication|Not required|EAC or Exchange Management Shell|
|OWA|Basic authentication|SSL required <br/> Requires 128-bit encryption|EAC or Exchange Management Shell|
|PowerShell|Anonymous authentication|Not required|Exchange Management Shell|
|Rpc|Basic authentication <br/> Windows authentication|SSL required <br/> Requires 128-bit encryption|Exchange Management Shell|
|RpcWithCert|By default, all authentication methods are disabled.|Required||
 
## Mailbox server

The following table lists the default settings on a stand-alone Exchange Mailbox server.

**Default Mailbox server IIS authentication and SSL settings**

|**Virtual directory**|**Authentication method**|**SSL settings**|**Management method**|
|:-----|:-----|:-----|:-----|
|Default website|Anonymous authentication|SSL required <br/> Requires 128-bit encryption|This virtual directory can't be configured by the user.|
|PowerShell|Anonymous authentication|Not required|Exchange Management Shell|
 
## See also

[Virtual directory management](https://technet.microsoft.com/library/ff952752(v=exchg.150).aspx)

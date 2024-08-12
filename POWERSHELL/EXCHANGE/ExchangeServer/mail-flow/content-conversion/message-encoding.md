---
title: "Message encoding options in Exchange Server"
ms.author: chrisda
author: chrisda
manager: serdars
ms.date: 7/6/2018
ms.audience: ITPro
ms.topic: article
ms.prod: exchange-server-it-pro
localization_priority: Normal
ms.assetid: c1d9edbb-d87c-41e5-881b-cd612d83d7e4
description: "Summary: Learn about the options that are available for message encoding in Exchange Server 2016 or Exchange Server 2019."
---

# Message encoding options in Exchange Server

The message encoding options in Exchange Server let you specify message characteristics such as MIME and non-MIME character sets, binary encoding, and attachment formats. You can specify message encoding options in the following locations:
  
- Remote domain settings
    
- Mail contact and mail user settings
    
- Outlook settings:
    
  - Message format
    
  - Internet message format
    
  - Internet recipient message format (Outlook 2010 or earlier)
    
  - Message character set encoding options
    
- Outlook on the web (formerly known as Outlook Web App) message format settings
    
Typically, the default settings for these message encoding options will work fine. However, you might need to change the messaging encoding options for recipients that are using older email clients or messaging systems. They'll likely tell you if messages from your Exchange environment appear to have formatting issues.
  
For more information about content conversion in Exchange, see [Content conversion](content-conversion.md). For TNEF (also known as or Rich Text) settings, see [TNEF conversion options](tnef-conversion.md).
  
## Remote domain settings
<a name="RemoteDomains"> </a>

Remote domains specify settings for messages sent to domains that are external to your Exchange organization. For more information, see [Remote Domains](http://technet.microsoft.com/library/10fb7d62-4d78-40a3-82db-d62bcd27ba42.aspx).
  
When you configure message encoding options for a remote domain, the settings are applied to all messages that are sent to recipients in that domain. Some settings are available in the Exchange admin center (EAC), but most are only available in the Exchange Management Shell. The message encoding settings are described in this table:
  
****

|**Setting** <br/> |**EAC configuration** <br/> |**Exchange Management Shell configuration** <br/> |
|:-----|:-----|:-----|
|**MIME character set**: The specified character set is only used for MIME messages that don't contain a character set. This setting won't overwrite character sets that are already specified in outgoing messages.  <br/> **Non-MIME character set**: This setting is used if either of these conditions are true:  <br/> • Incoming messages from a remote domain are missing the value of the _charset=_ setting in the MIME **Content-Type:** header field.  <br/> • Outgoing messages to a remote domain are missing the value of the MIME character set.  <br/> |**Mail flow** \> **Remote domains** \> **Add** ![Add icon](../../media/ITPro_EAC_AddIcon.png), or select an existing remote domain, and then click **Edit** ![Edit icon](../../media/ITPro_EAC_EditIcon.png) \> **Supported character set** section.  <br/> |Cmdlet: **Set-RemoteDomain** <br/> Parameters: _CharacterSet_ and _NonMimeCharacterSet_ <br/> |
|**Content type**: Valid values are:  <br/> `MimeHtmlText`: All messages are converted to MIME messages that use HTML formatting, unless the original message is a text message. If the original message is a text message, the outgoing message will be a MIME message that uses text formatting. This is the default value.  <br/> `MimeText`: All messages are converted to MIME messages that use text formatting.  <br/> `MimeHtml`: All messages are converted to MIME messages that use HTML formatting.  <br/> |n/a  <br/> |Cmdlet: **Set-RemoteDomain** <br/> Parameter: _ContentType_ <br/> |
|**Line wrap size**: You can specify the maximum number of characters that can exist on a single line of text in the body of the email message. Older email clients might prefer 78 characters per line.  <br/> |n/a  <br/> |Cmdlet: **Set-RemoteDomain** <br/> Parameter: _LineWrapSize_ <br/>  The default value is `Unlimited`, which means the email client is responsible for setting the line wrap size in new messages.  <br/> |
   
## Mail contact and mail user settings
<a name="Contacts"> </a>

Mail contacts and mail users represent users that have external email addresses in your Exchange organization. For more information, see [Recipients](../../recipients/recipients.md).
  
When you configure message encoding options for a mail contact or a mail user, the settings are only applied to messages that are sent to that specific recipient. All settings are only available in the Exchange Management Shell in these cmdlets:
  
- [Enable-MailContact](http://technet.microsoft.com/library/0accff85-3a03-4068-81e2-0508a4df21ec.aspx), [New-MailContact](http://technet.microsoft.com/library/c5abe0d4-3004-4d25-bda6-cb6155a47142.aspx), or [Set-MailContact](http://technet.microsoft.com/library/04c4e889-8546-4395-9d26-31af08264e45.aspx).
    
- [Enable-MailUser](http://technet.microsoft.com/library/1a6e86d0-09d8-4570-bf43-7ae6f1386c78.aspx), [New-MailUser](http://technet.microsoft.com/library/128467a7-b8b8-4fa6-bca9-1131301f18ce.aspx), or [Set-MailUser](http://technet.microsoft.com/library/087a55a2-ee8d-41a8-9c8f-d86e32ce8448.aspx).
    
The message encoding settings for mail contacts and mail users are described in this list:
  
- **UsePreferMessageFormat* parameter**: Specifies whether the message format settings for the mail contact or mail user override the corresponding settings for the remote domain. Valid values are:
    
  - `$true`: Messages sent to the Mail contact or mail user use the message format that's configured for the Mail contact or mail user.
    
  - `$false`: Messages sent to the Mail contact or mail user use the message format that's configured for the remote domain (the default remote domain or a specific remote domain) or configured by the message sender. This is the default value.
    
- **MessageFormat parameter**: This parameter specifies the message format for messages sent to the mail contact or mail user. Valid values are `Text` or `Mime`, and the default value is `Mime`.
    
- **MessageBodyFormat parameter**: This parameter specifies the message body format for messages sent to the mail contact or mail user. Valid values are `Text`, `Html`, or `TextAndHtml`, and the default value is `TextAndHtml`.
    
    The _MessageFormat_ and _MessageBodyFormat_ parameters are interdependent: 
    
  - If the _MessageFormat_ value is `Mime`, the _MessageBodyFormat_ value can be `Text`, `Html`, or `TextAndHtml`.
    
  - If the _MessageFormat_ value is `Text`, the _MessageBodyFormat_ value can only be `Text`.
    
- **MacAttachmentFormat parameter**: Specifies the message attachment format for Apple Macintosh operating system clients. Valid values are `BinHex`, `UuEncode`, `AppleSingle`, or `AppleDouble`, and the default value is `BinHex`.
    
    The _MessageFormat_ and _MacAttachmentFormat_ parameters are interdependent: 
    
  - If the _MessageFormat_ value is `Text`, the _MacAttachmentFormat_ value can be `BinHex` or `UuEncode`.
    
  - If the _MessageFormat_ value is `Mime`, the _MacAttachmentFormat_ value can be `BinHex`, `AppleSingle`, or `AppleDouble`.
    
## Outlook settings
<a name="Outlook"> </a>

As a sender, you can specify the message encoding in Outlook by using any of these methods:
  
- Configure the default message format to plain text or HTML.
    
- Configure the message format to plain text or HTML as you're composing the message by using the **Format** area in the **Format Text** tab.
    
- Configure the message encoding options for messages sent to all external recipients. These options are called *Internet message format* options, and they only apply to remote recipients (not to recipients in the Exchange organization).
    
- Configure the message encoding options for messages sent to specific external recipients (Outlook 2010 or earlier). These options are called *Internet recipient message format* options, and they only apply to remote recipients in your Contacts folder (not to recipients in the Exchange organization).
    
For instructions on configuring these settings in Outlook, see [Change the message format to HTML, Rich Text Format, or plain text](https://go.microsoft.com/fwlink/p/?linkid=397890).
  
By default, Outlook uses automatic character set message encoding by scanning the whole text of the outgoing message to determine the appropriate encoding to use for the message. This setting applies to internal and external recipients. However, you can bypass the automatic selection and specify a preferred encoding for outgoing messages at **File** \> **Options** \> **Advanced** \> **International options**.
  
## Outlook on the web settings
<a name="OutlookWebApp"> </a>

As a sender, you can specify message encoding options in Outlook on the web by using either of these methods:
  
- Configure the default message format as plain text or HTML in the **Message format** section at **Settings** \> **Options** \> **Mail** \> **Layout**.
    ![Options menu location in Outlook on the web](../../media/f1227a01-7f83-4af9-abf5-2c3dec6cf3d0.png)
  
- Configure the message format to plain text or HTML as you're composing the message by clicking **More options** ![More Options icon](../../media/ITPro_EAC_MoreOptionsIcon.png), and selecting **Switch to plain text** (if the current format is HTML) or **Switch to HTML** (if the current format is plain text).
    
## Order of precedence for message encoding options
<a name="Order"> </a>

Some message encoding options are available in remote domain settings, Mail contact or mail user settings, and Outlook or Outlook on the web settings. Message encoding options for outgoing messages sent to external recipients are described in the following list from highest priority to lowest priority:
  
1. Mail contact or mail user settings (if the use preferred message format setting is enabled)
    
2. Outlook or Outlook on the web settings
    
3. Remote domain settings
    
A setting at a higher level overrides the corresponding setting at a lower level. For example, Mail contact or mail user settings override the corresponding setting for a remote domain. Unique settings are unaffected (there's no higher or lower priority setting that conflicts).
  
The order of precedence for message encoding options are described in the following sections.
  
### Order of precedence for message character sets

The following table describes the order of precedence from highest priority to lowest priority for message character set encoding options.
  
|**Source**|**Setting**|**Values**|
|:-----|:-----|:-----|
|Outlook  <br/> |**Preferred encoding for outgoing messages** <br/> |**Automatically select encoding for outgoing messages** enabled or disabled (enabled by default).  <br/> **Preferred encoding for outgoing messages** set to the specified character set. This is the encoding option that's used if you disable **Automatically select encoding for outgoing messages** <br/> |
|Remote domain  <br/> |MIME character set and non-MIME character set  <br/> |The specified MIME and non-MIME character sets (which can be the same).  <br/> |
   
 **Notes**:
  
- When you configure the non-MIME character set for a remote domain, the character set is assigned to incoming or outgoing messages to and from the remote domain that don't contain a specified character set.
    
- The value of the Windows ANSI code page for the Exchange server is used to assign a character set to these types of messages:
    
  - Internal messages that don't contain a specified character set.
    
  - Internal messages that contain a specified character set, but don't contain a specified server code page.
    
- If a message contains a specified but invalid character set, the Exchange server tries to replace the invalid character set with a valid one.
    
### Order of precedence for plain text message encoding options

The following table describes the order of precedence from highest priority to lowest priority for plain text message encoding options.
  
 **Note**: Only plain text message settings are included here (not plain text settings for MIME encoded messages).
  
|**Source**|**Setting**|**Values**|
|:-----|:-----|:-----|
|Mail contact or mail user  <br/> |Use the preferred message format  <br/> |If the value `$true`, the plain text message encoding settings for the mail contact or mail user override the corresponding settings in Outlook.  <br/> If the value is `$false`, the plain text message encoding settings for the mail contact or mail user are ignored (the corresponding settings in Outlook are used).  <br/> |
|Mail contact or mail user  <br/> |Message format  <br/> |Text  <br/> |
|Mail contact or mail user  <br/> |Message body format  <br/> |Text  <br/> |
|Mail contact or mail user  <br/> |Mac attachment format  <br/> | `BinHex` or UUEncode  <br/> |
|Outlook 2010 or earlier  <br/> |Internet recipient message format (settings on a specific contact)  <br/> |Send plain text only  <br/> Open a contact in the Contacts folder \> double-click the email address \> click **View more options for interacting with this person** \> select **Outlook properties**, In the **E-mail Properties** dialog that opens, select **Send Plain Text only** in the **Internet format** field.  <br/> |
|Outlook  <br/> |Internet message format  <br/> |Plain text options for external messages at **File** \> **Options** \> **Mail** \> **Message format**:  <br/> **Encode attachments in UUENCODE format when sending plain-text messages** (not selected by default)  <br/> **Automatically wrap text at nn characters** (the default value is 76).  <br/> |
|Remote domain  <br/> |Line wrap size  <br/> |132 characters or less, or the value `Unlimited`. The default value is `Unlimited`.  <br/> |
   
### Order of precedence for MIME message encoding options

The following table describes the order of precedence from highest priority to lowest priority for MIME message encoding options.
  
|**Source**|**Setting**|**Values**|
|:-----|:-----|:-----|
|Mail contact or mail user  <br/> |Use the preferred message format  <br/> |If the value `$true`, the MIME message encoding settings for the mail contact or mail user override the corresponding settings in Outlook.  <br/> If the value is `$false`, the MIME text message encoding settings for the mail contact or mail user are ignored (the corresponding settings in Outlook, Outlook on the web, or remote domains are used).  <br/> |
|Mail contact or mail user  <br/> |Message format  <br/> |MIME  <br/> |
|Mail contact or mail user  <br/> |Message body format  <br/> |Text, HTML, or `TextAndHtml` (the default value is `TextAndHtml`).  <br/> |
|Mail contact or mail user  <br/> |Mac attachment format  <br/> | `BinHex`, `AppleSingle`, or `AppleDouble` (the default value is `BinHex`).  <br/> |
|Outlook or Outlook on the web  <br/> |Message format  <br/> |Plain text or HTML  <br/> |
|Remote domain  <br/> |Content type  <br/> | `MimeHtmlText` (the default value), `MimeText`, or `MimeHtml` <br/> |
   


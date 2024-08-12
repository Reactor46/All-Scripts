---
title: "Deploying Outlook for iOS and Android app configuration settings"
ms.author: dmaguire
author: msdmaguire
ms.reviewer: smithre4
manager: serdars
ms.date: 9/21/2018
ms.audience: ITPro
ms.topic: article
ms.service: exchange-online
localization_priority: Normal
ms.assetid: e8a034f6-39b8-4dea-a3bc-9421aaa75d1d
description: "Summary: How to customize the behavior of Outlook for iOS and Android in your Exchange organization."
---

# Deploying Outlook for iOS and Android app configuration settings

 **Summary**: How to customize the behavior of Outlook for iOS and Android in your Exchange organization.
  
Outlook for iOS and Android supports app settings that allow Office 365 and mobile device management (MDM), like Intune, administrators to customize the behavior of the app. 

Outlook for iOS and Android supports the following configuration scenarios:

- Account setup configuration
- Organization allowed accounts mode
- Data protection settings

Each configuration scenario will highlight its specific requirements; for example, whether the configuration scenario requires device enrollment, and thus work with any MDM provider, or requires Intune App Protection Policies.
 
## Account setup configuration settings
Outlook for iOS and Android offers administrators the ability to "push" account configurations to their Office 365 users. This capability only works with enrolled devices, however, it is supported with any MDM provider. If you are not using Intune, you will need to consult with your MDM documentation on how to deploy these settings. 

For more information on account setup configuration, see [Account setup with modern authentication in Exchange Online](https://docs.microsoft.com/exchange/clients-and-mobile-in-exchange-online/outlook-for-ios-and-android/setup-with-modern-authentication#account-setup-configuration-via-enterprise-mobility-management).

|**Key**|**Value**|**Device Enrollment Type**|
|:-----|:-----|:-----|
|com.microsoft.outlook.EmailProfile.EmailAddress|This value specifies the email address to be used for sending and receiving mail.  <br/> **Value type**: String  <br/> **Accepted values**: Email address  <br/> **Default if not specified**: \<blank\>  <br/> **Required**: Yes  <br/> **Example**: user@companyname.com  <br/> **Intune Token**<sup>*</sup>: {{mail}}|Managed devices|
|com.microsoft.outlook.EmailProfile.EmailUPN|This value specifies the User Principal Name or username for the email profile that will be used to authenticate the account.  <br/> **Value type**: String  <br/> **Accepted values**: UPN Address or username  <br/> **Default if not specified**: \<blank\>  <br/>**Required**: Yes  <br/> **Example**: userupn@companyname.com  <br/> **Intune Token**<sup>*</sup>: {{userprincipalname}}|Managed devices|
|com.microsoft.outlook.EmailProfile.AccountType|This value specifies the account type being configured based on the authentication model.  <br/> **Value type**: String  <br/> **Accepted values**: ModernAuth  <br/> **Required**: Yes  <br/> **Example**: ModernAuth|Managed devices|

## Organization allowed accounts mode settings
Outlook for iOS and Android offers administrators the ability to restrict email and storage provider accounts to only corporate accounts. This capability only works with enrolled devices, however, it is supported with any MDM provider. If you are not using Intune, you will need to consult with your MDM documentation on how to deploy these settings.

For more information on organization allowed accounts mode, please see [Account setup with modern authentication in Exchange Online](https://docs.microsoft.com/exchange/clients-and-mobile-in-exchange-online/outlook-for-ios-and-android/setup-with-modern-authentication#organization-allowed-accounts-mode).

|**Key**|**Value**|**Platform**|**Device Enrollment Type**|
|:-----|:-----|:-----|:-----|
|IntuneMAMAllowedAccountsOnly|This value specifies the whether organization allowed account mode is active.  <br/> **Value type**: String  <br/> **Accepted values**: Enabled, Disabled  <br/> **Required**: Yes  <br/> **Value**: Enabled|iOS|Managed devices|
|IntuneMAMUPN|This value specifies the User Principal Name for the account.  <br/> **Value type**: String  <br/> **Accepted values**: UPN Address  <br/> **Required**: Yes  <br/> **Example**: userupn@companyname.com  <br/> **Intune Token**<sup>*</sup>: {{userprincipalname}}|iOS|Managed devices|
|com.microsoft.intune.mam.AllowedAccountUPNs|This delimited value specifies the UPNs allowed for organization allowed account mode.  <br/> **Accepted values**: UPN Address  <br/> **Required**: Yes  <br/> **Example**: userupn@companyname.com  <br/> **Intune Token**<sup>*</sup>: {{userprincipalname}}|Android|Managed devices|

## Data protection settings
Outlook for iOS and Android supports app configuration policies for the following data protection settings when the app is managed by Intune:
- Managing the use of wearable technology
- Managing mail and calendar reminder notifications on iOS
- Managing the contact fields synced to the native contacts app

These settings can be deployed to the app regardless of device enrollment status.

### Configure Wearables for Outlook for iOS and Android

By default, Outlook for iOS and Android supports wearable technology, allowing the user to receive message notifications and event reminders, and the ability to interact with messages and view daily calendars. Organizations that want to disable the ability to access corporate data on wearables can deploy the following key via App configuration policies.
  
|**Key**|**Value**|**Device Enrollment Type**|
|:-----|:-----|:-----|
|com.microsoft.intune.mam.areWearablesAllowed|This value specifies if Outlook data can be synchronized to a wearable device. Setting the value to False disables wearable synchronization.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: False|Managed apps|
   
### Configure Notifications for Outlook for iOS

The Apple notification architecture ensures notifications are mirrored on iOS devices and WatchOS. Which device shows the notification depends on the device state: if the Apple Watch is unlocked and on a wrist, while the iOS device is locked, then WatchOS will alert the user with the notification. Apple does not provide a mechanism where you can administratively control and prevent notifications on WatchOS while still allowing them to be delivered on iOS devices. 

The following configuration settings will disable notifications completely on iOS and WatchOS. The disadvantage is that the end user will never see new mail notifications or calendar reminders on iOS devices. The user will have to launch the Outlook for iOS in order to discover new mail or see calendar appointments.
 
|**Key**|**Value**|**Device Enrollment Type**|
|:-----|:-----|:-----|
|com.microsoft.outlook.Mail.NotificationsEnabled|This value specifies if Outlook will allow mail notifications. Setting the value to False disables mail notifications.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: False|Managed apps|
|com.microsoft.outlook.Mail.NotificationsEnabled.UserChangeAllowed|This value specifies if the user can adjust the mail notification setting within the app. Setting the value to False prevents the user from adjusting the mail notification setting.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: False|Managed apps|
|com.microsoft.outlook.Calendar.NotificationsEnabled|This value specifies if Outlook will allow calendar reminder notifications. Setting the value to False disables calendar reminder notifications.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: False|Managed apps|
|com.microsoft.outlook.Calendar.NotificationsEnabled.UserChangeAllowed|This value specifies if the user can adjust the calendar reminder notification setting within the app. Setting the value to False prevents the user from adjusting the calendar reminder notification setting.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: False|Managed apps|

### Configure Contact Field Sync to native Contacts for Outlook for iOS and Android
<a name="contacts"> </a>

The settings in the following table allow you to control the contact fields that will sync between Outlook on iOS and Android and the native Contacts applications.
  
> [!NOTE]
> Outlook for Android supports bi-directional contact synchronization. If a user edits a field in the native contacts app that is restricted (such as the **Notes** field), then that data will not synchronize back into Outlook for Android. 
  
|**Key**|**Value**|**Device Enrollment Type**|
|:-----|:-----|:-----|
|com.microsoft.outlook.ContactSync.AddressAllowed|This value specifies if the contact's address should be synchronized to native contacts.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: True|Managed apps|
|com.microsoft.outlook.ContactSync.BirthdayAllowed|This value specifies if the contact's birthday should be synchronized to native contacts.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: True|Managed apps|
|com.microsoft.outlook.ContactSync.CompanyAllowed|This value specifies if the contact's company name should be synchronized to native contacts.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: True|Managed apps|
|com.microsoft.outlook.ContactSync.DepartmentAllowed|This value specifies if the contact's department should be synchronized to native contacts.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: True|Managed apps|
|com.microsoft.outlook.ContactSync.EmailAllowed|This value specifies if the contact's email address should be synchronized to native contacts.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: True|Managed apps|
|com.microsoft.outlook.ContactSync.InstantMessageAllowed|This value specifies if the contact's instant messaging address should be synchronized to native contacts.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: True|Managed apps|
|com.microsoft.outlook.ContactSync.JobTitleAllowed|This value specifies if the contact's job title should be synchronized to native contacts.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: True|Managed apps|
|com.microsoft.outlook.ContactSync.NicknameAllowed|This value specifies if the contact's nickname should be synchronized to native contacts.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: True|Managed apps|
|com.microsoft.outlook.ContactSync.NotesAllowed|This value specifies if the contact's notes should be synchronized to native contacts.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: True|Managed apps|
|com.microsoft.outlook.ContactSync.PhoneHomeAllowed|This value specifies if the contact's home phone number should be synchronized to native contacts.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: True|Managed apps|
|com.microsoft.outlook.ContactSync.PhoneHomeFaxAllowed|This value specifies if the contact's home fax number should be synchronized to native contacts.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: True|Managed apps|
|com.microsoft.outlook.ContactSync.PhoneMobileAllowed|This value specifies if the contact's mobile phone number should be synchronized to native contacts.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: True|Managed apps|
|com.microsoft.outlook.ContactSync.PhoneOtherAllowed|This value specifies if the contact's other phone number should be synchronized to native contacts.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: True|Managed apps|
|com.microsoft.outlook.ContactSync.PhonePagerAllowed|This value specifies if the contact's pager phone number should be synchronized to native contacts.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: True|Managed apps|
|com.microsoft.outlook.ContactSync.PhoneWorkAllowed|This value specifies if the work phone number should be synchronized to native contacts.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: True|Managed apps|
|com.microsoft.outlook.ContactSync.PhoneWorkFaxAllowed|This value specifies if the contact's work fax number should be synchronized to native contacts.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: True|Managed apps|
|com.microsoft.outlook.ContactSync.PrefixAllowed|This value specifies if the contact's name prefix should be synchronized to native contacts.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: True|Managed apps|
|com.microsoft.outlook.ContactSync.SuffixAllowed|This value specifies if the contact's name suffix should be synchronized to native contacts.  <br/> **Accepted values**: True, False  <br/> **Default if not specified**: True  <br/> **Example**: True|Managed apps|
   
## Deploying the configuration scenarios with Microsoft Intune

If you are using Microsoft Intune as your mobile device management provider, the following steps will allow you to create an app configuration policy. After the configuration is created, you can assign its settings to groups of users.

> [!NOTE]
> Intune managed apps will check-in with an interval of 30 minutes for Intune App Configuration Policy status, when deployed in conjunction with an Intune App Protection Policy. If an Intune App Protection Policy isn't assigned to the user, then the Intune App Configuration Policy check-in interval is set to 720 minutes.

> [!IMPORTANT]
> When deploying app configuration policies to managed devices, issues can occur when multiple policies have different values for the same configuration key and are targeted for the same app and user. This is due to the lack of a conflict resolution mechanism for resolving the differing values. You can prevent this by ensuring only a single app configuration policy for managed devices is defined and targeted for the same app and user. 

### Create an app configuration policy for Outlook for iOS and Android
  
1. Sign in to the Azure portal.
    
2. Select **More Services** \> **Monitoring + Management** \> **Intune**.
    
3. On the **Client apps** blade of the Manage list, select **App configuration policies**.
    
4. On the **App Configuration policies** blade, choose **Add**. 
    
5. On the **Add app configuration** blade, enter a **Name**, and optional **Description** for the app configuration settings. 
    
6. For **Device enrollment** type, choose **Managed apps** or **Managed devices** depending on the configuration scenario you are deploying (see the specific configuration scenario for more information).
    
7. If you chose **Managed devices** you will have another option, **Platform**. For **Platform**, choose either **iOS** or **Android**.

8. For **Associated app**, choose **Select the required app**, and then, on the **Targeted apps** blade, choose **Outlook**. If you specified, **Managed apps**, select both the iOS and Android platform Outlook apps. 
    
    > [!NOTE]
    > If Outlook is not listed as an available app, then you must add it by following the instructions in [Add Android store apps to Microsoft Intune](https://docs.microsoft.com/intune/store-apps-android) and [Add iOS store apps to Microsoft Intune](https://docs.microsoft.com/intune/store-apps-ios). 
  
9. Click **OK** to return to the **Add app configuration** blade. 
    
10. Choose **Configuration Settings**. On the **Configuration** blade, define the key and value pairs that will supply configurations for Outlook for iOS and Android. The key and value pairs you can define are described in the following sections. 
    
11. When you are done, choose **OK**.
    
12. On the **Add app configuration** blade, choose **Add**.
    
The newly created configuration will be displayed on the **App configuration** blade. 

>[!NOTE]
> If you selected **Managed devices**, you will need to create a separate app configuration policy for each platform.
  
### Assign the configuration settings that you created

You assign the settings to groups of users in Azure Active Directory. When a user has the Microsoft Outlook app installed, the app will be managed by the settings you have specified. To do this:
  
1. From the **Intune** blade, on the **Mobile apps** blade of the Manage list, select **App configuration policies**.
    
2. From the list of app configuration policies, select the one you want to assign.
    
3. On the next blade, choose **Assignments**.
    
4. On the **Assignments** blade, select the Azure AD group to which you want to assign the app configuration, and then choose **OK**.

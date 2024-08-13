---
title: "Managing Outlook for iOS and Android in Exchange Online"
ms.author: dmaguire
author: msdmaguire
ms.reviewer: smithre4
manager: laurawi
ms.date: 3/30/2018
ms.audience: ITPro
ms.topic: article
ms.service: exchange-online
localization_priority: Normal
ms.assetid: 30e8c819-4d41-4458-a746-a9ba9a84a7c0
description: "Summary: This article describes best practices for managing mobile devices with Outlook for iOS and Android in Exchange Online."
---

# Managing Outlook for iOS and Android in Exchange Online

 **Summary**: This article describes best practices for managing mobile devices with Outlook for iOS and Android in Exchange Online.
  
Outlook for iOS and Android provides users the fast, intuitive email and calendar experience users expect from a modern mobile app, while being the only app to provide support for the best features of Office 365. In addition, Microsoft provides a number of utilities for managing and protecting company data on mobile devices in your Exchange Online organization.
  
## Options for managing devices and applications in Office 365

Customers looking to manage Outlook for iOS and Android have the following options: 
  
1. **Recommended**: The Enterprise Mobility + Security suite, which includes Microsoft Intune and Azure Active Directory conditional access.
    
2. Mobile Device Management (MDM) for Office 365.
    
3. Mobile Device Access and Mobile Device Mailbox Policies.
    
> [!NOTE]
> For implementation details on each of these three options, see [Securing Outlook for iOS and Android in Exchange Online](secure-outlook-for-ios-and-android.md). 
  
Microsoft recommends Office 365 customers use the features of the Enterprise Mobility + Security suite to protect corporate data on mobile devices, due to the advanced capabilities provided by these services. The core capabilities of the built-in MDM for Office 365 are included with an Office 365 subscription, while the broader capabilities of the Enterprise Mobility + Security require an additional subscription purchase.
  
> [!IMPORTANT]
> Mobile device access rules (allow, block, or quarantine) in Exchange Online are skipped when access is managed by a conditional access policy that includes either [Require device to be marked as compliant](https://docs.microsoft.com/azure/active-directory/active-directory-conditional-access-policy-connected-applications) or [Require approved client app](https://docs.microsoft.com/azure/active-directory/active-directory-conditional-access-technical-reference). 
  
A complete side-by-side comparison of MDM and Intune is available in [Choose between MDM for Office 365 and Microsoft Intune](https://support.office.com/article/c93d9ab9-efb2-4349-9b93-30c30562ee22).
  
> [!NOTE]
> When using mobile device cmdlets such as `Get-MobileDevice` to check the status of a device, the timestamp for Outlook for iOS and Android synchronization, indicated by the `LastSyncTime` property, may be up to 15 minutes behind the actual time of synchronization. While device synchronization does occur in real time, the returned time stamp may lag behind. 
  
### Using Enterprise Mobility + Security

The richest and broadest protection capabilities for Office 365 data are available when you subscribe to the Enterprise Mobility + Security suite, which includes Microsoft Intune, Azure Information Protection, and Azure Active Directory Premium features, such as conditional access.
  
> [!NOTE]
> While the Enterprise Mobility + Security suite subscription includes licenses for both Microsoft Intune and Azure Active Directory, customers can purchase Microsoft Intune licenses and Azure Active Directory Premium licenses separately. All users must be licensed to leverage the conditional access and Intune app protection policies discussed in this article. 
  
Intune provides mobile application management (MAM) capabilities, as well as other conditional access and device management capabilities. With Intune app protection policies, you can restrict actions such as cut, copy, paste, and "save as" of corporate data between Intune-managed apps and apps that are not managed by Intune. More information is available in [How to create and assign app protection policies](https://docs.microsoft.com/intune/app-protection-policies). Additionally, the Intune-managed Outlook apps include a new multi-identity management feature that enables users to access both their personal and work email accounts in the same Outlook app while only applying the Intune app protection policies to the user's work account. This provides a much more seamless user experience.
  
Conditional access is a capability of Azure Active Directory that enables you to enforce controls on the access to apps in your environment based on specific conditions from a central location. By using conditional access policies, you can apply the right access controls under the required conditions. Azure Active Directory conditional access provides you with added security when such security is needed, and it stays out of your users' way when it isn't.
  
Key features of the Enterprise Mobility + Security suite with Outlook for iOS and Android:
  
- **Conditional access**. Azure Active Directory ensures that Exchange Online email can be accessed only when the conditional access requirements are met. For more information on device enrollment, see [Conditional access in Azure Active Directory](https://docs.microsoft.com/azure/active-directory/active-directory-conditional-access-azure-portal).
    
- **Intune app protection**. Outlook for iOS and Android allows you to protect your corporate data with Intune app protection policies. This is a great option for "bring your own device" (BYOD) scenarios where you want to keep corporate data safe without managing a users' devices. For more information on Intune app protection policies, see [Protect app data using mobile app management policies with Microsoft Intune](https://docs.microsoft.com/intune/deploy-use/protect-app-data-using-mobile-app-management-policies-with-microsoft-intune).
    
- **Device enrollment**. Intune lets you manage your workforce's devices and apps, and how they access your company data. In this model, Outlook for iOS and Android ensures that Exchange Online email can be accessed only on phones and tablets that are managed by your company and are compliant with your organization's policy. When users log on to the Outlook app on an unmanaged mobile device, Outlook prompts users to enroll the device in Intune by leveraging the Azure conditional access policy, and then validates that the device meets organizational standards of device compliance.
    
- **Device management and reporting**. The enrollment process allows organizations to set and manage security policies that, for example, enforce device-level PIN lock, require data encryption, and block compromised devices in order to prevent untrusted devices from accessing corporate email and data. Each enrolled device appears in the Office 365 admin center, and reporting is available to provide details on the devices that access your corporate data.
    
- **Selective wipe**. Microsoft Intune can remove Office 365 email data from Outlook for iOS and Android, while leaving any personal email accounts intact (whether the device is enrolled or not). This is an increasingly important requirement as more businesses adopt a "bring your own device" approach to phones and tablets.
    
 For more about Microsoft Intune see [Documentation for Microsoft Intune](https://docs.microsoft.com/intune/).
  
### Using built-in Mobile Device Management (MDM) for Office 365

MDM for Office 365 provides device management capabilities at no additional cost. Microsoft Intune powers these basic capabilities, providing a core set of controls in the Office 365 admin center for organizations that need the basics.
  
Because this is a device management solution, there is no native capability to control which apps can be used, even after a device is enrolled. If you want to limit access to Outlook for iOS and Android, you will need to obtain Azure Active Directory Premium licenses and leverage conditional access policies.
  
Outlook for iOS and Android fully supports the capabilities provided by MDM for Office 365.
  
For detailed information on MDM, see the following resources:
  
- [Overview built-in Mobile Device Management for Office 365](https://go.microsoft.com/fwlink/p/?LinkId=623837).
    
- [Manage settings and features on your devices with Microsoft Intune policies](https://docs.microsoft.com/intune/deploy-use/manage-settings-and-features-on-your-devices-with-microsoft-intune-policies)
    
- Instructions for your end-users to enroll a device in Office 365 MDM: [Enroll your mobile device in Office 365](https://support.office.com/article/c8ac722d-dcaf-4135-8345-3e6327f5d3c5)
    
### Using Mobile Device Access and Mobile Device Mailbox Policies

Microsoft recommends Office 365 customers use either the Enterprise Mobility + Security suite or the built-in MDM for Office 365 to manage company data on mobile devices, due to the advanced capabilities provided by those services. Outlook for iOS and Android does support mobile device access and mobile device mailbox policies (formerly known as Exchange Active Sync policies), which are available through the Exchange admin center.
  
Outlook for iOS and Android supports the following Exchange mobile device mailbox policy settings:
  
- Device encryption enabled
    
- Min password length
    
- Password enabled
    
See [Mobile device mailbox policies in Exchange Online](../../clients-and-mobile-in-exchange-online/exchange-activesync/mobile-device-mailbox-policies.md) for more information. 
  
Exchange administrators can initiate a remote device wipe against Outlook for iOS and Android. Upon receiving the remote wipe request, the app will remove the profile and all data associated with it.
  


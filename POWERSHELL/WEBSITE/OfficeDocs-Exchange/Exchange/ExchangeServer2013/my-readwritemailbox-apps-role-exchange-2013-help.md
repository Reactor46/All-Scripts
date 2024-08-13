﻿---
title: 'My ReadWriteMailbox Apps role: Exchange 2013 Help'
TOCTitle: My ReadWriteMailbox Apps role
ms:assetid: febb73fb-3a0b-4c67-b53b-9566d7c32cd2
ms:mtpsurl: https://technet.microsoft.com/en-us/library/Mt793556(v=EXCHG.150)
ms:contentKeyID: 74418060
ms.date: 02/08/2017
mtps_version: v=EXCHG.150
---

# My ReadWriteMailbox Apps role

 


The My ReadWriteMailbox Apps role enables individual users to add add-ins that request the ReadWriteMailbox permission in their manifest.

This management role is one of several built-in roles in the Role Based Access Control (RBAC) permissions model in Microsoft Exchange Server 2013. Management roles, which are assigned to one or more management role groups, management role assignment policies, users, or universal security groups (USG), act as a logical grouping of cmdlets or scripts that are combined to provide access to view or modify the configuration of Exchange 2013 components, such as mailbox databases, transport rules, and recipients. If a cmdlet or script and its parameters, together called a management role entry, are included on a role, that cmdlet or script and its parameters can be run by those assigned the role. For more information about management roles and management role entries, see [Understanding management roles](understanding-management-roles-exchange-2013-help.md).

For more information about management roles, management role groups, and other RBAC components, see [Understanding Role Based Access Control](understanding-role-based-access-control-exchange-2013-help.md).

## Management role assignments

For this role to grant permissions, it must be assigned to a role assignee, such as a role assignment policy. This assignment is done using management role assignments. Role assignments link role assignees and roles together. If more than one role is assigned to a role assignee, the role assignee is granted the combination of all the permissions granted by all the assigned roles.


> [!NOTE]
> You can also assign this management role to a role group, USG, or directly to a user. However user-focused roles are most effective when used with role assignment policies.



This user-focused role has implicit scopes that can't be modified. Therefore, you shouldn't add custom scopes to role assignments that assign this role to role assignment policies, role groups, USGs, or users.

For more information about role assignments and scopes, see the following topics:

  - [Understanding management role assignments](understanding-management-role-assignments-exchange-2013-help.md)

  - [Understanding management role scopes](understanding-management-role-scopes-exchange-2013-help.md)

This role may be assigned to one or more role assignment policies by default. For more information, see the "Default Management Role Assignments" section.

If you want to view a list of role groups, users, or USGs assigned to this role, use the following command.

```powershell
Get-ManagementRoleAssignment -Role "<role name>"
```

## Regular and delegating role assignments

This role can be assigned to role assignees using either regular or delegating role assignments. Regular role assignments grant the permissions provided by the role to the role assignee. Delegating role assignments grant the role assignee the ability to assign the role to other role assignees. For more information about regular and delegating role assignments, see [Understanding management role assignments](understanding-management-role-assignments-exchange-2013-help.md).

## Adding or removing role assignments

You can change which role assignees are assigned this role. By changing which role assignee is assigned this role, you change who is granted its permissions. You can assign this role to other built-in role assignment policies, or you can create role assignment policies and assign this role to them.

To assign this role to role assignees, the role must be assigned to a role group you're a member of, directly to you, or to a USG you're a member of, using a delegating role assignment. For more information about delegating role assignments, see the "Regular and Delegating Role Assignments" section.

You can also remove this role from the default role assignment policy, role assignment policies and role groups you create, users, and USGs. However, there must always be at least one delegating role assignment between this role and a role group or USG. You can't delete the last delegating role assignment.

For more information about how to add or remove assignments between this role and role groups, users, and USGs, see the following topics:

  - "Add or remove a role to or from a role group" section in [Manage role groups](manage-role-groups-exchange-2013-help.md)

  - [Add a role to a user or USG](add-a-role-to-a-user-or-usg-exchange-2013-help.md)

  - [Remove a role from a user or USG](remove-a-role-from-a-user-or-usg-exchange-2013-help.md)

## Enabling or disabling role assignments

By enabling or disabling a role assignment, you control whether that role assignment should be in effect. If a role assignment is disabled, the permissions granted by the associated role aren't applied to the role assignee. This is convenient if you want to temporarily remove permissions without deleting a role assignment. For more information, see [Change a role assignment](change-a-role-assignment-exchange-2013-help.md).

## Default management role assignments

This role has role assignments to one or more role assignees. The following table indicates whether the role assignment is regular or delegating, and also indicates the management scopes applied to each assignment. The following list describes each column:

  - **Regular assignment**   Regular role assignments enable the role assignee to access the permissions provided by the management role entries on this role.

  - **Delegating assignment**   Delegating role assignments give the role assignee the ability to assign this role to role groups, users, or USGs.

  - **Recipient read scope**   The recipient read scope determines what recipient objects the role assignee is allowed to read from Active Directory.

  - **Recipient write scope**   The recipient write scope determines what recipient objects the role assignee is allowed to modify in Active Directory.

  - **Configuration read scope**   The configuration read scope determines what configuration and server objects the role assignee is allowed to read from Active Directory.

  - **Configuration write scope**   The configuration write scope determines what organizational and server objects the role assignee is allowed to modify in Active Directory.

### Default management role assignments for this role

<table style="width:100%;">
<colgroup>
<col style="width: 14%" />
<col style="width: 14%" />
<col style="width: 14%" />
<col style="width: 14%" />
<col style="width: 14%" />
<col style="width: 14%" />
<col style="width: 14%" />
</colgroup>
<thead>
<tr class="header">
<th>Role group or assignment policy</th>
<th>Regular assignment</th>
<th>Delegating assignment</th>
<th>Recipient read scope</th>
<th>Recipient write scope</th>
<th>Configuration read scope</th>
<th>Configuration write scope</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><p><a href="organization-management-exchange-2013-help.md">Organization Management</a></p></td>
<td><p> </p></td>
<td><p>X</p></td>
<td><p><code>Self</code></p></td>
<td><p><code>Self</code></p></td>
<td><p><code>OrganizationConfig</code></p></td>
<td><p><code>OrganizationConfig</code></p></td>
</tr>
</tbody>
</table>


## Management role customization

This role has been configured to provide a role assignee with all necessary cmdlets and parameters to manage the features and components listed in the beginning of this topic. Other roles have also been provided to enable management of other features. By adding and removing roles to and from role groups, you can create a customized permissions model without the need to customize individual management roles. For a complete list of roles, see [Built-in management roles](built-in-management-roles-exchange-2013-help.md). For more information about customizing role groups, see [Manage role groups](manage-role-groups-exchange-2013-help.md).

If you decide that you need to create a customized version of this role, you must create a role as a child of this role, and customize the new role.


> [!WARNING]
> The following information enables you to perform advanced management of permissions. Customizing management roles can significantly increase the complexity of your permissions model. You could cause certain features to stop functioning if you replace a built-in management role with an incorrectly configured custom role.



The following are the most common steps to create a customized role and assign it to a role assignee:

1.  Create a copy of this role. For more information, see [Create a role](create-a-role-exchange-2013-help.md).

2.  Change or remove the role entries on the new role using the **Set-ManagementRoleEntry** and **Remove-ManagementRoleEntry** cmdlets. You can't add additional role entries to the new role because it can only contain the role entries on the parent built-in role. For more information, see the following topics:
    
      - [Change a role entry](change-a-role-entry-exchange-2013-help.md)
    
      - [Remove a role entry from a role](remove-a-role-entry-from-a-role-exchange-2013-help.md)

3.  If you want to replace the built-in role with this new customized role, remove any role assignments associated with the built-in role. For more information, see the following topics:
    
      - “Add or remove a role to or from a role group” section in [Manage role groups](manage-role-groups-exchange-2013-help.md)
    
      - [Remove a role from a user or USG](remove-a-role-from-a-user-or-usg-exchange-2013-help.md)

4.  Add the new customized role to the required role assignees. For more information, see the following topics:
    
      - “Add or remove a role to or from a role group” section in [Manage role groups](manage-role-groups-exchange-2013-help.md)
    
      - [Add a role to a user or USG](add-a-role-to-a-user-or-usg-exchange-2013-help.md)
        

        > [!IMPORTANT]
        > If you want other users, in addition to the user that created the role, to be able to assign the new customized role, be sure to add a delegating role assignment to at least one role assignee. For more information, see <A href="delegate-role-assignments-exchange-2013-help.md">Delegate role assignments</A>.



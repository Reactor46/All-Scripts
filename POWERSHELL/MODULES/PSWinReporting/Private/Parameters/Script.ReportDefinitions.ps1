## Define reports
$Script:ReportDefinitions = @{
    UserChanges             = @{
        Enabled = $false
        Events  = @{
            Enabled     = $false
            Events      = 4720, 4738
            LogName     = 'Security'
            Fields      = [ordered] @{
                'Computer'            = 'Domain Controller'
                'Action'              = 'Action'
                'ObjectAffected'      = 'User Affected'
                'SamAccountName'      = 'SamAccountName'
                'DisplayName'         = 'DisplayName'
                'UserPrincipalName'   = 'UserPrincipalName'
                'HomeDirectory'       = 'Home Directory'
                'HomePath'            = 'Home Path'
                'ScriptPath'          = 'Script Path'
                'ProfilePath'         = 'Profile Path'
                'UserWorkstations'    = 'User Workstations'
                'PasswordLastSet'     = 'Password Last Set'
                'AccountExpires'      = 'Account Expires'
                'PrimaryGroupId'      = 'Primary Group Id'
                'AllowedToDelegateTo' = 'Allowed To Delegate To'
                'OldUacValue'         = 'Old Uac Value'
                'NewUacValue'         = 'New Uac Value'
                'UserAccountControl'  = 'User Account Control'
                'UserParameters'      = 'User Parameters'
                'SidHistory'          = 'Sid History'
                'Who'                 = 'Who'
                'Date'                = 'When'
                # Common Fields
                'ID'                  = 'Event ID'
                'RecordID'            = 'Record ID'
                'GatheredFrom'        = 'Gathered From'
                'GatheredLogName'     = 'Gathered LogName'
            }
            Ignore      = @{
                # Cleanup Anonymous LOGON (usually related to password events)
                # https://social.technet.microsoft.com/Forums/en-US/5b2a93f7-7101-43c1-ab53-3a51b2e05693/eventid-4738-user-account-was-changed-by-anonymous?forum=winserverDS
                SubjectUserName = "ANONYMOUS LOGON"

                # Test value
                #ProfilePath     = 'C*'
            }
            Functions   = @{
                'ProfilePath'        = 'Convert-UAC'
                'OldUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                'NewUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                'UserAccountControl' = 'Remove-WhiteSpace', 'Split-OnSpace', 'Convert-UAC'
            }
            IgnoreWords = @{
                #'Profile Path' = 'TEMP*'
            }
            SortBy      = 'When'
        }
    }
    UserChangesDetailed     = [ordered] @{
        Enabled = $false
        Events  = @{
            Enabled     = $false
            Events      = 5136, 5137, 5141
            LogName     = 'Security'
            Filter      = @{
                # Filter is special, if there is just one object on the right side
                # If there are more objects filter will pick all values on the right side and display them as required
                'ObjectClass' = 'user'
                #'OperationType'            = 'Value Added'
                #'AttributeLDAPDisplayName' = $null,'displayName' #, 'versionNumber'
            }
            Functions   = @{
                #'ProfilePath'        = 'Convert-UAC'
                #'OldUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                #'NewUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                #'UserAccountControl' = 'Remove-WhiteSpace', 'Split-OnSpace', 'Convert-UAC'
                #'DSType'        = 'ConvertFrom-OperationType'
                'OperationType' = 'ConvertFrom-OperationType'
            }
            Fields      = [ordered] @{

                'Computer'                 = 'Domain Controller'
                'Action'                   = 'Action'
                'OperationType'            = 'Action Detail'
                'Who'                      = 'Who'
                'Date'                     = 'When'


                'ObjectDN'                 = 'User Object'
                #'ObjectGUID'               = 'ObjectGUID'
                #'ObjectClass'              = 'ObjectClass'
                'AttributeLDAPDisplayName' = 'Field Changed'
                #'AttributeSyntaxOID'       = 'AttributeSyntaxOID'
                'AttributeValue'           = 'Field Value'

                #'OpCorrelationID'          = 'OperationCorelationID'
                #'AppCorrelationID'         = 'OperationApplicationCorrelationID'

                #'DSName'                   = 'DSName'
                #'DSType'                   = 'DSType'
                #'Task'                     = 'Task'
                #'Version'                  = 'Version'

                # Common Fields
                'RecordID'                 = 'Record ID'
                'ID'                       = 'Event ID'
                'GatheredFrom'             = 'Gathered From'
                'GatheredLogName'          = 'Gathered LogName'
            }

            SortBy      = 'Record ID'
            Descending  = $false
            IgnoreWords = @{

            }
        }
    }
    ComputerChangesDetailed = [ordered] @{
        Enabled = $false
        Events  = @{
            Enabled     = $false
            Events      = 5136, 5137, 5141
            LogName     = 'Security'
            Filter      = @{
                # Filter is special, if there is just one object on the right side
                # If there are more objects filter will pick all values on the right side and display them as required
                'ObjectClass' = 'computer'
                #'OperationType'            = 'Value Added'
                #'AttributeLDAPDisplayName' = $null,'displayName' #, 'versionNumber'
            }
            Functions   = @{
                #'ProfilePath'        = 'Convert-UAC'
                #'OldUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                #'NewUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                #'UserAccountControl' = 'Remove-WhiteSpace', 'Split-OnSpace', 'Convert-UAC'
                #'DSType'        = 'ConvertFrom-OperationType'
                'OperationType' = 'ConvertFrom-OperationType'
            }
            Fields      = [ordered] @{

                'Computer'                 = 'Domain Controller'
                'Action'                   = 'Action'
                'OperationType'            = 'Action Detail'
                'Who'                      = 'Who'
                'Date'                     = 'When'


                'ObjectDN'                 = 'Computer Object'
                #'ObjectGUID'               = 'ObjectGUID'
                #'ObjectClass'              = 'ObjectClass'
                'AttributeLDAPDisplayName' = 'Field Changed'
                #'AttributeSyntaxOID'       = 'AttributeSyntaxOID'
                'AttributeValue'           = 'Field Value'

                #'OpCorrelationID'          = 'OperationCorelationID'
                #'AppCorrelationID'         = 'OperationApplicationCorrelationID'

                #'DSName'                   = 'DSName'
                #'DSType'                   = 'DSType'
                #'Task'                     = 'Task'
                #'Version'                  = 'Version'

                # Common Fields
                'RecordID'                 = 'Record ID'
                'ID'                       = 'Event ID'
                'GatheredFrom'             = 'Gathered From'
                'GatheredLogName'          = 'Gathered LogName'
            }

            SortBy      = 'Record ID'
            Descending  = $false
            IgnoreWords = @{

            }
        }
    }
    UserStatus              = @{
        Enabled = $false
        Events  = @{
            Enabled     = $false
            Events      = 4722, 4725, 4767, 4723, 4724, 4726
            LogName     = 'Security'
            IgnoreWords = @{}
            Fields      = [ordered] @{
                'Computer'        = 'Domain Controller'
                'Action'          = 'Action'
                'Who'             = 'Who'
                'Date'            = 'When'
                'ObjectAffected'  = 'User Affected'

                # Common Fields
                'ID'              = 'Event ID'
                'RecordID'        = 'Record ID'
                'GatheredFrom'    = 'Gathered From'
                'GatheredLogName' = 'Gathered LogName'
            }
            SortBy      = 'When'
        }
    }
    UserLockouts            = @{
        Enabled = $false
        Events  = @{
            Enabled     = $false
            Events      = 4740
            LogName     = 'Security'
            IgnoreWords = @{}
            Fields      = [ordered] @{
                'Computer'         = 'Domain Controller'
                'Action'           = 'Action'
                'TargetDomainName' = 'Computer Lockout On'
                'ObjectAffected'   = 'User Affected'
                'Who'              = 'Reported By'
                'Date'             = 'When'

                # Common Fields
                'ID'               = 'Event ID'
                'RecordID'         = 'Record ID'
                'GatheredFrom'     = 'Gathered From'
                'GatheredLogName'  = 'Gathered LogName'
            }
            SortBy      = 'When'
        }
    }
    UserLogon               = @{
        Enabled = $false
        Events  = @{
            Enabled     = $false
            Events      = 4624
            LogName     = 'Security'
            Fields      = [ordered] @{
                'Computer'           = 'Computer'
                'Action'             = 'Action'
                'IpAddress'          = 'IpAddress'
                'IpPort'             = 'IpPort'
                'ObjectAffected'     = 'User / Computer Affected'
                'Who'                = 'Who'
                'Date'               = 'When'
                'LogonProcessName'   = 'LogonProcessName'
                'ImpersonationLevel' = 'ImpersonationLevel' # %%1833 = Impersonation
                'VirtualAccount'     = 'VirtualAccount'  #  %%1843 = No
                'ElevatedToken'      = 'ElevatedToken' # %%1842 = Yes
                'LogonType'          = 'LogonType'
                # Common Fields
                'ID'                 = 'Event ID'
                'RecordID'           = 'Record ID'
                'GatheredFrom'       = 'Gathered From'
                'GatheredLogName'    = 'Gathered LogName'
            }
            IgnoreWords = @{}
        }
    }
    UserUnlocked            = @{
        # 4767	A user account was unlocked
        # https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventid=4767
        Enabled = $false
        Events  = @{
            Enabled     = $false
            Events      = 4767
            LogName     = 'Security'
            IgnoreWords = @{}
            Functions   = @{
                #'ProfilePath'        = 'Convert-UAC'
                #'OldUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                #'NewUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                #'UserAccountControl' = 'Remove-WhiteSpace', 'Split-OnSpace', 'Convert-UAC'
                #'DSType'        = 'ConvertFrom-OperationType'
                # 'IpAddress' = 'Clean-IpAddress'
            }

            Fields      = [ordered] @{
                'Computer'         = 'Domain Controller'
                'Action'           = 'Action'
                'TargetDomainName' = 'Computer Lockout On'
                'ObjectAffected'   = 'User Affected'
                'Who'              = 'Who'
                'Date'             = 'When'

                # Common Fields
                'ID'               = 'Event ID'
                'RecordID'         = 'Record ID'
                'GatheredFrom'     = 'Gathered From'
                'GatheredLogName'  = 'Gathered LogName'
            }

            SortBy      = 'When'
        }
    }
    ComputerCreatedChanged  = @{
        Enabled = $false
        Events  = @{
            Enabled     = $false
            Events      = 4741, 4742 # created, changed
            LogName     = 'Security'
            Ignore      = @{
                # Cleanup Anonymous LOGON (usually related to password events)
                # https://social.technet.microsoft.com/Forums/en-US/5b2a93f7-7101-43c1-ab53-3a51b2e05693/eventid-4738-user-account-was-changed-by-anonymous?forum=winserverDS
                SubjectUserName = "ANONYMOUS LOGON"
            }
            Fields      = [ordered] @{
                'Computer'            = 'Domain Controller'
                'Action'              = 'Action'
                'ObjectAffected'      = 'Computer Affected'
                'SamAccountName'      = 'SamAccountName'
                'DisplayName'         = 'DisplayName'
                'UserPrincipalName'   = 'UserPrincipalName'
                'HomeDirectory'       = 'Home Directory'
                'HomePath'            = 'Home Path'
                'ScriptPath'          = 'Script Path'
                'ProfilePath'         = 'Profile Path'
                'UserWorkstations'    = 'User Workstations'
                'PasswordLastSet'     = 'Password Last Set'
                'AccountExpires'      = 'Account Expires'
                'PrimaryGroupId'      = 'Primary Group Id'
                'AllowedToDelegateTo' = 'Allowed To Delegate To'
                'OldUacValue'         = 'Old Uac Value'
                'NewUacValue'         = 'New Uac Value'
                'UserAccountControl'  = 'User Account Control'
                'UserParameters'      = 'User Parameters'
                'SidHistory'          = 'Sid History'
                'Who'                 = 'Who'
                'Date'                = 'When'
                # Common Fields
                'ID'                  = 'Event ID'
                'RecordID'            = 'Record ID'
                'GatheredFrom'        = 'Gathered From'
                'GatheredLogName'     = 'Gathered LogName'
            }
            IgnoreWords = @{}
        }
    }
    ComputerDeleted         = @{
        Enabled = $false
        Events  = @{
            Enabled     = $false
            Events      = 4743 # deleted
            LogName     = 'Security'
            IgnoreWords = @{}
            Fields      = [ordered] @{
                'Computer'        = 'Domain Controller'
                'Action'          = 'Action'
                'ObjectAffected'  = 'Computer Affected'
                'Who'             = 'Who'
                'Date'            = 'When'

                # Common Fields
                'ID'              = 'Event ID'
                'RecordID'        = 'Record ID'
                'GatheredFrom'    = 'Gathered From'
                'GatheredLogName' = 'Gathered LogName'
            }
            SortBy      = 'When'
        }
    }
    UserLogonKerberos       = @{
        Enabled = $false
        Events  = @{
            Enabled     = $false
            Events      = 4768
            LogName     = 'Security'
            IgnoreWords = @{}
            Functions   = @{
                #'ProfilePath'        = 'Convert-UAC'
                #'OldUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                #'NewUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                #'UserAccountControl' = 'Remove-WhiteSpace', 'Split-OnSpace', 'Convert-UAC'
                #'DSType'        = 'ConvertFrom-OperationType'
                'IpAddress' = 'Clean-IpAddress'
            }
            Fields      = [ordered] @{
                'Computer'             = 'Domain Controller'
                'Action'               = 'Action'
                'ObjectAffected'       = 'Computer/User Affected'
                'IpAddress'            = 'IpAddress'
                'IpPort'               = 'Port'
                'TicketOptions'        = 'TicketOptions'
                'Status'               = 'Status'
                'TicketEncryptionType' = 'TicketEncryptionType'
                'PreAuthType'          = 'PreAuthType'
                'Date'                 = 'When'

                # Common Fields
                'ID'                   = 'Event ID'
                'RecordID'             = 'Record ID'
                'GatheredFrom'         = 'Gathered From'
                'GatheredLogName'      = 'Gathered LogName'
            }
            SortBy      = 'When'
        }
    }
    GroupMembershipChanges  = @{
        Enabled = $false
        Events  = @{
            Enabled     = $false
            Events      = 4728, 4729, 4732, 4733, 4746, 4747, 4751, 4752, 4756, 4757, 4761, 4762, 4785, 4786, 4787, 4788
            LogName     = 'Security'
            IgnoreWords = @{
                'Who' = '*ANONYMOUS*'
            }
            Fields      = [ordered] @{
                'Computer'            = 'Domain Controller'
                'Action'              = 'Action'
                'TargetUserName'      = 'Group Name'
                'MemberNameWithoutCN' = 'Member Name' # Required work {$_.MemberName -replace '^CN=|,.*$' }}, fixed in PSEventViewer
                'Who'                 = 'Who'
                'Date'                = 'When'

                # Common Fields
                'ID'                  = 'Event ID'
                'RecordID'            = 'Record ID'
                'GatheredFrom'        = 'Gathered From'
                'GatheredLogName'     = 'Gathered LogName'
            }
            SortBy      = 'When'
        }
    }
    GroupEnumeration        = @{
        Enabled = $false
        Events  = @{
            Enabled     = $false
            Events      = 4798, 4799
            LogName     = 'Security'
            IgnoreWords = @{
                'Who' = '*ANONYMOUS*'
            }
            Fields      = [ordered] @{
                'Computer'        = 'Domain Controller'
                'Action'          = 'Action'
                'TargetUserName'  = 'Group Name'
                #'MemberNameWithoutCN' = 'Member Name' # Required work {$_.MemberName -replace '^CN=|,.*$' }}, fixed in PSEventViewer
                'Who'             = 'Who'
                'Date'            = 'When'

                # Common Fields
                'ID'              = 'Event ID'
                'RecordID'        = 'Record ID'
                'GatheredFrom'    = 'Gathered From'
                'GatheredLogName' = 'Gathered LogName'
            }
            SortBy      = 'When'
        }
    }
    GroupChanges            = @{
        Enabled = $false
        Events  = @{
            Enabled     = $false
            Events      = 4735, 4737, 4745, 4750, 4760, 4764, 4784, 4791
            LogName     = 'Security'
            IgnoreWords = @{
                'Who' = '*ANONYMOUS*'
            }

            Fields      = [ordered] @{
                'Computer'        = 'Domain Controller'
                'Action'          = 'Action'
                'TargetUserName'  = 'Group Name'
                'Who'             = 'Who'
                'Date'            = 'When'
                'GroupTypeChange' = 'Changed Group Type'
                'SamAccountName'  = 'Changed SamAccountName'
                'SidHistory'      = 'Changed SidHistory'

                # Common Fields
                'ID'              = 'Event ID'
                'RecordID'        = 'Record ID'
                'GatheredFrom'    = 'Gathered From'
                'GatheredLogName' = 'Gathered LogName'
            }

            SortBy      = 'When'
        }
    }
    GroupCreateDelete       = @{
        Enabled = $false
        Events  = @{
            Enabled     = $false
            Events      = 4727, 4730, 4731, 4734, 4744, 4748, 4749, 4753, 4754, 4758, 4759, 4763
            LogName     = 'Security'
            IgnoreWords = @{
                'Who' = '*ANONYMOUS*'
            }
            Fields      = [ordered] @{
                'Computer'        = 'Domain Controller'
                'Action'          = 'Action'
                'TargetUserName'  = 'Group Name'
                'Who'             = 'Who'
                'Date'            = 'When'
                #'GroupTypeChange' = 'Group Type Change'

                # Common Fields
                'ID'              = 'Event ID'
                'RecordID'        = 'Record ID'
                'GatheredFrom'    = 'Gathered From'
                'GatheredLogName' = 'Gathered LogName'
            }
            SortBy      = 'When'
        }
    }
    GroupChangesDetailed = [ordered] @{
        Enabled = $false
        Events  = @{
            Enabled     = $false
            Events      = 5136, 5137, 5141
            LogName     = 'Security'
            Filter      = @{
                # Filter is special, if there is just one object on the right side
                # If there are more objects filter will pick all values on the right side and display them as required
                'ObjectClass' = 'group'
                #'OperationType'            = 'Value Added'
                #'AttributeLDAPDisplayName' = $null,'displayName' #, 'versionNumber'
            }
            Functions   = @{
                #'ProfilePath'        = 'Convert-UAC'
                #'OldUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                #'NewUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                #'UserAccountControl' = 'Remove-WhiteSpace', 'Split-OnSpace', 'Convert-UAC'
                #'DSType'        = 'ConvertFrom-OperationType'
                'OperationType' = 'ConvertFrom-OperationType'
            }
            Fields      = [ordered] @{

                'Computer'                 = 'Domain Controller'
                'Action'                   = 'Action'
                'OperationType'            = 'Action Detail'
                'Who'                      = 'Who'
                'Date'                     = 'When'


                'ObjectDN'                 = 'Computer Object'
                #'ObjectGUID'               = 'ObjectGUID'
                'ObjectClass'              = 'ObjectClass'
                'AttributeLDAPDisplayName' = 'Field Changed'
                #'AttributeSyntaxOID'       = 'AttributeSyntaxOID'
                'AttributeValue'           = 'Field Value'

                #'OpCorrelationID'          = 'OperationCorelationID'
                #'AppCorrelationID'         = 'OperationApplicationCorrelationID'

                #'DSName'                   = 'DSName'
                #'DSType'                   = 'DSType'
                #'Task'                     = 'Task'
                #'Version'                  = 'Version'

                # Common Fields
                'RecordID'                 = 'Record ID'
                'ID'                       = 'Event ID'
                'GatheredFrom'             = 'Gathered From'
                'GatheredLogName'          = 'Gathered LogName'
            }

            SortBy      = 'Record ID'
            Descending  = $false
            IgnoreWords = @{

            }
        }
    }
    GroupPolicyChanges      = [ordered] @{
        Enabled                     = $false
        'Group Policy Name Changes' = @{
            Enabled     = $false
            Events      = 5136, 5137, 5141
            LogName     = 'Security'
            Filter      = @{
                # Filter is special, if there is just one object on the right side
                # If there are more objects filter will pick all values on the right side and display them as required
                'ObjectClass'              = 'groupPolicyContainer'
                #'OperationType'            = 'Value Added'
                'AttributeLDAPDisplayName' = $null, 'displayName' #, 'versionNumber'
            }
            Functions   = @{
                #'ProfilePath'        = 'Convert-UAC'
                #'OldUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                #'NewUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                #'UserAccountControl' = 'Remove-WhiteSpace', 'Split-OnSpace', 'Convert-UAC'
                #'DSType'        = 'ConvertFrom-OperationType'
                'OperationType' = 'ConvertFrom-OperationType'
            }
            Fields      = [ordered] @{
                'RecordID'                 = 'Record ID'
                'Computer'                 = 'Domain Controller'
                'Action'                   = 'Action'
                'Who'                      = 'Who'
                'Date'                     = 'When'


                'ObjectDN'                 = 'ObjectDN'
                'ObjectGUID'               = 'ObjectGUID'
                'ObjectClass'              = 'ObjectClass'
                'AttributeLDAPDisplayName' = 'AttributeLDAPDisplayName'
                #'AttributeSyntaxOID'       = 'AttributeSyntaxOID'
                'AttributeValue'           = 'AttributeValue'
                'OperationType'            = 'OperationType'
                'OpCorrelationID'          = 'OperationCorelationID'
                'AppCorrelationID'         = 'OperationApplicationCorrelationID'

                'DSName'                   = 'DSName'
                'DSType'                   = 'DSType'
                'Task'                     = 'Task'
                'Version'                  = 'Version'

                # Common Fields
                'ID'                       = 'Event ID'

                'GatheredFrom'             = 'Gathered From'
                'GatheredLogName'          = 'Gathered LogName'
            }

            SortBy      = 'Record ID'
            Descending  = $false
            IgnoreWords = @{

            }
        }
        'Group Policy Edits'        = @{
            Enabled     = $false
            Events      = 5136, 5137, 5141
            LogName     = 'Security'
            Filter      = @{
                # Filter is special, if there is just one object on the right side
                # If there are more objects filter will pick all values on the right side and display them as required
                'ObjectClass'              = 'groupPolicyContainer'
                #'OperationType'            = 'Value Added'
                'AttributeLDAPDisplayName' = 'versionNumber'
            }
            Functions   = @{
                #'ProfilePath'        = 'Convert-UAC'
                #'OldUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                #'NewUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                #'UserAccountControl' = 'Remove-WhiteSpace', 'Split-OnSpace', 'Convert-UAC'
                #'DSType'        = 'ConvertFrom-OperationType'
                'OperationType' = 'ConvertFrom-OperationType'
            }
            Fields      = [ordered] @{
                'RecordID'                 = 'Record ID'
                'Computer'                 = 'Domain Controller'
                'Action'                   = 'Action'
                'Who'                      = 'Who'
                'Date'                     = 'When'


                'ObjectDN'                 = 'ObjectDN'
                'ObjectGUID'               = 'ObjectGUID'
                'ObjectClass'              = 'ObjectClass'
                'AttributeLDAPDisplayName' = 'AttributeLDAPDisplayName'
                #'AttributeSyntaxOID'       = 'AttributeSyntaxOID'
                'AttributeValue'           = 'AttributeValue'
                'OperationType'            = 'OperationType'
                'OpCorrelationID'          = 'OperationCorelationID'
                'AppCorrelationID'         = 'OperationApplicationCorrelationID'

                'DSName'                   = 'DSName'
                'DSType'                   = 'DSType'
                'Task'                     = 'Task'
                'Version'                  = 'Version'

                # Common Fields
                'ID'                       = 'Event ID'

                'GatheredFrom'             = 'Gathered From'
                'GatheredLogName'          = 'Gathered LogName'
            }

            SortBy      = 'Record ID'
            Descending  = $false
            IgnoreWords = @{

            }
        }
        'Group Policy Links'        = @{
            Enabled     = $false
            Events      = 5136, 5137, 5141
            LogName     = 'Security'
            Filter      = @{
                # Filter is special, if there is just one object on the right side
                # If there are more objects filter will pick all values on the right side and display them as required
                'ObjectClass' = 'domainDNS'
                #'OperationType'            = 'Value Added'
                #'AttributeLDAPDisplayName' = 'versionNumber'
            }
            Functions   = @{
                #'ProfilePath'        = 'Convert-UAC'
                #'OldUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                #'NewUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                #'UserAccountControl' = 'Remove-WhiteSpace', 'Split-OnSpace', 'Convert-UAC'
                #'DSType'        = 'ConvertFrom-OperationType'
                'OperationType' = 'ConvertFrom-OperationType'
            }
            Fields      = [ordered] @{
                'RecordID'                 = 'Record ID'
                'Computer'                 = 'Domain Controller'
                'Action'                   = 'Action'
                'Who'                      = 'Who'
                'Date'                     = 'When'


                'ObjectDN'                 = 'ObjectDN'
                'ObjectGUID'               = 'ObjectGUID'
                'ObjectClass'              = 'ObjectClass'
                'AttributeLDAPDisplayName' = 'AttributeLDAPDisplayName'
                #'AttributeSyntaxOID'       = 'AttributeSyntaxOID'
                'AttributeValue'           = 'AttributeValue'
                'OperationType'            = 'OperationType'
                'OpCorrelationID'          = 'OperationCorelationID'
                'AppCorrelationID'         = 'OperationApplicationCorrelationID'

                'DSName'                   = 'DSName'
                'DSType'                   = 'DSType'
                'Task'                     = 'Task'
                'Version'                  = 'Version'

                # Common Fields
                'ID'                       = 'Event ID'

                'GatheredFrom'             = 'Gathered From'
                'GatheredLogName'          = 'Gathered LogName'
            }

            SortBy      = 'Record ID'
            Descending  = $false
            IgnoreWords = @{

            }
        }
        <#

        Events2     = @{
            Enabled     = $false
            Events      = 5136, 5137, 5141
            LogName     = 'Security'
            Filter      = @{
                # Filter is special, if there is just one object on the right side
                # If there are more objects filter will pick all values on the right side and display them as required
                'ObjectClass'   = 'groupPolicyContainer'
                'OperationType' = 'Value Added'
                # 'AttributeLDAPDisplayName' = 'displayName','versionNumber'
            }
            Functions   = @{
                #'ProfilePath'        = 'Convert-UAC'
                #'OldUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                #'NewUacValue'        = 'Remove-WhiteSpace', 'Convert-UAC'
                #'UserAccountControl' = 'Remove-WhiteSpace', 'Split-OnSpace', 'Convert-UAC'
                #'DSType'        = 'ConvertFrom-OperationType'
                'OperationType' = 'ConvertFrom-OperationType'
            }
            Fields      = [ordered] @{


                'Computer'                 = 'Domain Controller'
                'Action'                   = 'Action'
                'Who'                      = 'Who'
                'Date'                     = 'When'


                'ObjectDN'                 = 'OBjectDN'
                'ObjectGUID'               = 'ObjectGUID'
                'ObjectClass'              = 'ObjectClass'
                'AttributeLDAPDisplayName' = 'AttributeLDAPDisplayName'
                #'AttributeSyntaxOID'       = 'AttributeSyntaxOID'
                'AttributeValue'           = 'AttributeValue'
                'OperationType'            = 'OperationType'

                'OpCorrelationID'          = 'OperationCorelationID'
                'AppCorrelationID'         = 'OperationApplicationCorrelationID'

                'DSName'                   = 'DSName'
                'DSType'                   = 'DSType'
                'Task'                     = 'Task'
                'Version'                  = 'Version'

                # Common Fields
                'ID'                       = 'Event ID'
                'RecordID'                 = 'Record ID'
                'GatheredFrom'             = 'Gathered From'
                'GatheredLogName'          = 'Gathered LogName'
            }

            SortBy      = 'Record ID'
            Descending  = $false
            IgnoreWords = @{

            }
        }
        EventsType3 = @{

        }
#>
    }
    LogsClearedSecurity     = @{
        Enabled = $false
        Events  = @{
            Enabled     = $false
            Events      = 1102, 1105
            LogName     = 'Security'

            Fields      = [ordered] @{
                'Computer'        = 'Domain Controller'
                'Action'          = 'Action'
                'BackupPath'      = 'Backup Path'
                'Channel'         = 'Log Type'

                'Who'             = 'Who'
                'Date'            = 'When'

                # Common Fields
                'ID'              = 'Event ID'
                'RecordID'        = 'Record ID'
                'GatheredFrom'    = 'Gathered From'
                'GatheredLogName' = 'Gathered LogName'
                #'Test' = 'Test'
            }
            SortBy      = 'When'
            IgnoreWords = @{}
            Overwrite   = @{
                # Allows to overwrite field content on the fly, either only on IF or IF ELSE
                # IF <VALUE> -eq <VALUE> THEN <VALUE> (3 VALUES)
                # IF <VALUE> -eq <VALUE> THEN <VALUE> ELSE <VALUE> (4 VALUES)
                # If you need to use IF multiple times for same field use spaces to distinguish HashTable Key.

                'Backup Path' = 'Backup Path', '', 'N/A'
                #'Backup Path ' = 'Backup Path', 'C:\Windows\System32\Winevt\Logs\Archive-Security-2018-11-24-09-25-36-988.evtx', 'MMMM'
                'Who'         = 'Event ID', 1105, 'Automatic Backup'  # if event id 1105 set field to Automatic Backup
                #'Test' = 'Event ID', 1106, 'Test', 'Mama mia'
            }
        }
    }
    LogsClearedOther        = @{
        Enabled = $false
        Events  = @{
            Enabled     = $false
            Events      = 104
            LogName     = 'System'
            IgnoreWords = @{}
            Fields      = [ordered] @{
                'Computer'     = 'Domain Controller'
                'Action'       = 'Action'
                'BackupPath'   = 'Backup Path'
                'Channel'      = 'Log Type'

                'Who'          = 'Who'
                'Date'         = 'When'

                # Common Fields
                'ID'           = 'Event ID'
                'RecordID'     = 'Record ID'
                'GatheredFrom' = 'Gathered From'
            }
            SortBy      = 'When'
            Overwrite   = @{
                # Allows to overwrite field content on the fly, either only on IF or IF ELSE
                # IF <VALUE> -eq <VALUE> THEN <VALUE> (3 VALUES)
                # IF <VALUE> -eq <VALUE> THEN <VALUE> ELSE <VALUE> (4 VALUES)
                # If you need to use IF multiple times for same field use spaces to distinguish HashTable Key.

                'Backup Path' = 'Backup Path', '', 'N/A'
            }
        }
    }
    EventsReboots           = @{
        Enabled = $false
        Events  = @{
            Enabled     = $false
            Events      = 1001, 1018, 1, 12, 13, 42, 41, 109, 1, 6005, 6006, 6008, 6013
            LogName     = 'System'
            IgnoreWords = @{

            }
            #SortBy      = 'When'
        }
    }
}
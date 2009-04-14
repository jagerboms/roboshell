insert into dbo.helpObjects (SystemID, ObjectName, HelpText, ColourText, State, AuditID)
select  'default', x.ObjectName, x.HelpText, x.ColourText, 'ac', 1
from
(
    select  ObjectName = 'shlSecurityAdmin'
           ,HelpText = 'The Users form is used to display accounts that have permission to access the system. The data is 
displayed using a standard Grid Form. All the associated features of the grid form are supported.<br />
User access is controlled within the Server using a tool like Microsoft SQL Server Management Studio. This screen provides a view of this information.'
           ,ColourText = 'The colours used are determined by the Type column as follows:'
    union select
            'shlUserPerm'
           ,'The Security Administration form is used to display and maintain account permissions. The data is displayed using a standard Tree Form. All of the associated features of the tree form are supported.
The application is defined as a collection of modules. The modules are organised into child parent relationships where children are displayed on a branch off the parent. When a permission is given to a module it is automatically inherited by all it children and their children etc unless the child has a specific permission of it own. This would allow for example an account to be granted permission to the top most module (displayed as the account name on the screen) and thus have access to the entire application.
Permissions can be granted, revoked or denied. Grant and deny permissions are implemented in the same manner the corresponding commands would in the database. The revoke permission is used to break the inheritance chain that in fact indicates the account is not granted or denied permission to the module and it''s children. A deny permission overrides a grant permission, so if a user belongs to more than one group and one group has denied permission on a module then the user is denied access even if another group he is a member of has access.
The delete action allows a permission to be removed and the module''s permission is then determined by it''s parents permission via inheritance.'
           ,ColourText = 'Colours are used to show inherited permissions as follows:'
    union select
            'shlUserPerm'
           ,'The Security Administration form is used to display and maintain account permissions. The data is displayed using a standard Tree Form. All of the associated features of the tree form are supported.
The application is defined as a collection of modules. The modules are organised into child parent relationships where children are displayed on a branch off the parent. When a permission is given to a module it is automatically inherited by all it children and their children etc unless the child has a specific permission of it own. This would allow for example an account to be granted permission to the top most module (displayed as the account name on the screen) and thus have access to the entire application.
Permissions can be granted, revoked or denied. Grant and deny permissions are implemented in the same manner the corresponding commands would in the database. The revoke permission is used to break the inheritance chain that in fact indicates the account is not granted or denied permission to the module and it''s children. A deny permission overrides a grant permission, so if a user belongs to more than one group and one group has denied permission on a module then the user is denied access even if another group he is a member of has access.
The delete action allows a permission to be removed and the module''s permission is then determined by it''s parents permission via inheritance.'
           ,ColourText = 'Colours are used to show inherited permissions as follows:'
    union select
            'shlRoleMembers'
           ,'The Role Members form is used to display members of a database role. The data is displayed using a standard Grid Form. All of the associated features of the grid form are supported.<br />
Database roles are created and maintained in SQL Server using a tool like Enterprise Manager. This information is provided to assist manage application security management and displays information taken from system tables in the application database.'
           ,null
    union select
            'shlGroupMember'
           ,'The Members form is used to display members of a windows security group. The data is displayed using a standard Grid Form. All of the associated features of the grid form are supported.<br />
The members of the requested group are displayed.<br />
The group name and description are displayed in the caption of the form.'
           ,null
) x
left join dbo.helpObjects o
on      o.SystemID = 'default'
and     o.ObjectName = x.ObjectName
where   o.SystemID is null
go

insert into dbo.helpFields (SystemID, ObjectName, FieldName, HelpText, State, AuditID)
select  'default', x.ObjectName, x.FieldName, x.HelpText, 'ac', 1
from
(
    select  ObjectName = 'shlSecurityAdmin'
           ,FieldName = 'UserName'
           ,HelpText = 'Name of the accessing account.Name of the accessing account.'
    union select
            'shlSecurityAdmin'
           ,'Type'
           ,'The type of account. A user can a SQL User, SQL Role, Windows User or Windows Group. The value of this column determines the colour used to display the row.'
    union select
            'shlSecurityAdmin'
           ,'Permissioned'
           ,'A flag (Yes/No) indicating whether or not the account has direct permissions in the security system. Direct permission does not include permissions obtained via membership to a role or group.'
    union select
            'shlUserPerm'
           ,'ModuleID'
           ,'The name of the application module.'
    union select
            'shlRoleMembers'
           ,'MemberName'
           ,'The account name of the role member.'
    union select
            'shlGroupMember'
           ,'MemberName'
           ,'The user name of the active directory group member.'
) x
left join dbo.helpFields f
on      f.SystemID = 'default'
and     f.ObjectName = x.ObjectName
and     f.FieldName = x.FieldName
where   f.SystemID is null
go

insert into dbo.helpActions (SystemID, ObjectName, ActionName, HelpText, State, AuditID)
select  'default', x.ObjectName, x.ActionName, x.HelpText, 'ac', 1
from
(
    select  ObjectName = 'shlSecurityAdmin'
           ,ActionName = 'Refresh'
           ,HelpText = 'Refresh the displayed information by re-reading the database.'
    union select
            'shlSecurityAdmin'
           ,'Permissions'
           ,'Displays and allows the maintenance of module permissions for the selected account.'
    union select
            'shlSecurityAdmin'
           ,'Members'
           ,'Displays details of the members of a group, either for <a href="RoleMembers.html">Database Role</a> or <a href="GroupMembers.html">Windows Group</a>.'
    union select
            'shlUserPerm'
           ,'Refresh'
           ,'Refresh the displayed information by re-reading the database.'
    union select
            'shlUserPerm'
           ,'Grant'
           ,'Allows access to granted to the module and it''s children. When selected the user is prompted "Do you wish to grant this permission?" clicking Yes sets the permission while clicking No leaves the setting unchanged.'
    union select
            'shlUserPerm'
           ,'Revoke'
           ,'Stops the module and it''s children inheriting permissions from it''s parent. When selected the user is prompted "Do you wish to revoke this permission?" clicking Yes sets the permission while clicking No leaves the setting unchanged.'
    union select
            'shlUserPerm'
           ,'Deny'
           ,'Denies access for the module and it''s children. When selected the user is prompted "Do you wish to deny this permission?" clicking Yes sets the permission while clicking No leaves the setting unchanged.'
    union select
            'shlUserPerm'
           ,'Delete'
           ,'Allows a previous permission setting to be removed. The module then inherits it''s parent''s permission. When selected the user is prompted "Do you wish to remove this permission rule?" clicking Yes deletes the permission while clicking No leaves the setting unchanged.'
    union select
            'shlRoleMembers'
           ,'Refresh'
           ,'Refresh the displayed information by re-reading the database.'
    union select
            'shlGroupMember'
           ,'Refresh'
           ,'Refresh the displayed information by re-reading the database.'
    union select
            'shlGroupMember'
           ,'Exit'
           ,'Exits the form.'
) x
left join dbo.helpActions a
on      a.SystemID = 'default'
and     a.ObjectName = x.ObjectName
and     a.ActionName = x.ActionName
where   a.SystemID is null
go

insert into dbo.helpColours (SystemID, ObjectName, ColourValue, ValueDescription, State, AuditID)
select  'default', x.ObjectName, x.ColourValue, x.ValueDescription, 'ac', 1
from
(
    select  ObjectName = 'shlSecurityAdmin'
           ,ColourValue = ''
           ,ValueDescription = 'Windows User'
    union select
            ObjectName = 'shlSecurityAdmin'
           ,ColourValue = 'G'
           ,ValueDescription = 'Windows Group'
    union select
            ObjectName = 'shlSecurityAdmin'
           ,ColourValue = 'R'
           ,ValueDescription = 'Database Role'
    union select
            ObjectName = 'shlSecurityAdmin'
           ,ColourValue = 'S'
           ,ValueDescription = 'SQL User'
    union select
            ObjectName = 'shlUserPerm'
           ,ColourValue = ''
           ,ValueDescription = 'Access not granted. (Revoked)'
    union select
            ObjectName = 'shlUserPerm'
           ,ColourValue = 'GA'
           ,ValueDescription = 'Access allowed. (Granted)'
    union select
            ObjectName = 'shlUserPerm'
           ,ColourValue = 'DA'
           ,ValueDescription = 'Access not allowed. (Denied)'
    union select
            ObjectName = 'shlUserPerm'
           ,ColourValue = 'GI'
           ,ValueDescription = ''
    union select
            ObjectName = 'shlUserPerm'
           ,ColourValue = 'DI'
           ,ValueDescription = ''
) x
left join dbo.helpColours c
on      c.SystemID = 'default'
and     c.ObjectName = x.ObjectName
and     c.ColourValue = x.ColourValue
where   c.SystemID is null
go

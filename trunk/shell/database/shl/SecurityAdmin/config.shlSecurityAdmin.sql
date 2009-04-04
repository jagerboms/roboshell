print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlGridFormInsert
    @ObjectName = 'shlSecurityAdmin'
   ,@Title = 'Users'
   ,@DataParameter = 'shlSecurityAdmin'
   ,@ColourColumn = 'Role'
   ,@HelpPage = 'Users.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'shlSecurityAdmin'
   ,@ModuleID = 'security'
   ,@ObjectName = 'shlSecurityAdmin'
go

---------------------------------------------------

execute dbo.shlFieldParamInsert
    @ObjectName = 'shlSecurityAdmin'
   ,@FieldName = 'UserName'
   ,@Label = 'User'
   ,@Width = 128
   ,@DisplayWidth = 200
   ,@ValueType = 'String'
   ,@IsPrimary = 'Y'
   ,@IsInput = 'N'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'shlSecurityAdmin'
   ,@FieldName = 'LoginName'
   ,@ValueType = 'String'
   ,@Width = 128
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
go

execute dbo.shlFieldsInsert
    @ObjectName = 'shlSecurityAdmin'
   ,@FieldName = 'Type'
   ,@Label = 'Type'
   ,@ValueType = 'String'
   ,@Width = 5
   ,@DisplayWidth = 100
go

execute dbo.shlFieldsInsert
    @ObjectName = 'shlSecurityAdmin'
   ,@FieldName = 'Permissioned'
   ,@Label = 'Permissioned'
   ,@ValueType = 'String'
   ,@Width = 3
   ,@DisplayWidth = 60
go

execute dbo.shlFieldsInsert
    @ObjectName = 'shlSecurityAdmin'
   ,@FieldName = 'Role'
   ,@Label = 'Role'
   ,@ValueType = 'String'
   ,@Width = 1
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlSecurityAdmin'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'R'
   ,@Value = 'Green'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlSecurityAdmin'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'G'
   ,@Value = 'Orange'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlSecurityAdmin'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'S'
   ,@Value = 'Black'
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'shlSecurityAdmin'
   ,@ActionName = 'Refresh'
   ,@Process = 'shlSecurityAdminGet'
   ,@ImageFile = 'Refresh.gif'
   ,@ToolTip = 'Refresh user details'
   ,@KeyCode = 116
go

execute dbo.shlActionsInsert
    @ObjectName = 'shlSecurityAdmin'
   ,@ActionName = 'Permissions'
   ,@Process = 'shlUserPerm'
   ,@RowBased = 'Y'
   ,@ImageFile = 'Permission.gif'
   ,@ToolTip = 'User permissions'
go

execute dbo.shlActionsInsert
    @ObjectName = 'shlSecurityAdmin'
   ,@ActionName = 'Members'
   ,@RowBased = 'Y'
   ,@ImageFile = 'Roles.gif'
   ,@ToolTip = 'Display members'
   ,@ProcessField = 'Role'
go

execute dbo.shlActionProcessRulesInsert
    @ObjectName = 'shlSecurityAdmin'
   ,@ActionName = 'Members'
   ,@Value = 'R'
   ,@Process = 'shlRoleMembers'
go

execute dbo.shlActionProcessRulesInsert
    @ObjectName = 'shlSecurityAdmin'
   ,@ActionName = 'Members'
   ,@Value = 'G'
   ,@Process = 'shlGroupMember'
go

-- execute dbo.shlActionsInsert
--     @ObjectName = 'shlSecurityAdmin'
--    ,@ActionName = 'Report'
--    ,@Process = 'psecUserReportGrid'
--    ,@ImageFile = 'Report.gif'
--    ,@ToolTip = 'User report'
-- go

print '.oOo.'
go

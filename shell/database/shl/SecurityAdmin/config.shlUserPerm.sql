print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlTreeInsert
    @ObjectName = 'shlUserPerm'
   ,@Title = 'Permissions'
   ,@KeyColumn = 'ModuleID'
   ,@DescriptionColumn = 'Description'
   ,@ParentColumn = 'Parent'
   ,@TypeColumn = 'Type'
   ,@ColourColumn = 'Type'
   ,@DefaultImage = 'None.gif'
   ,@TitleParameters = 'UserName'
   ,@RefreshTree = 'Y'
   ,@HelpPage = 'Permissions.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'shlUserPerm'
   ,@ModuleID = 'security'
   ,@ObjectName = 'shlUserPerm'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'shlUserPerm'
   ,@ParameterName = 'UserName'
   ,@ValueType = 'String'
   ,@Width = 128
go

execute dbo.shlParametersInsert
    @ObjectName = 'shlUserPerm'
   ,@ParameterName = 'ModuleID'
   ,@ValueType = 'String'
   ,@Width = 32
   ,@IsInput = 'N'
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlUserPerm'
   ,@PropertyType = 'im'
   ,@PropertyName = 'GA'
   ,@Value = 'Grant.gif'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlUserPerm'
   ,@PropertyType = 'im'
   ,@PropertyName = 'DA'
   ,@Value = 'Deny.gif'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlUserPerm'
   ,@PropertyType = 'im'
   ,@PropertyName = 'NA'
   ,@Value = 'NoneA.gif'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlUserPerm'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'GA'
   ,@Value = 'Green'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlUserPerm'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'GI'
   ,@Value = 'Green'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlUserPerm'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'DA'
   ,@Value = 'Red'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlUserPerm'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'DI'
   ,@Value = 'Red'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlUserPerm'
   ,@PropertyType = 'lk'
   ,@PropertyName = 'SECADMINTREE'
   ,@Value = ''
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'shlUserPerm'
   ,@ActionName = 'Refresh'
   ,@Process = 'shlUserPermGet'
   ,@ImageFile = 'Refresh.gif'
   ,@ToolTip = 'Refresh user permissions'
   ,@KeyCode = 116
go

execute dbo.shlActionsInsert
    @ObjectName = 'shlUserPerm'
   ,@ActionName = 'Grant'
   ,@Process = 'shlModuleUserGrant'
   ,@RowBased = 'Y'
   ,@ImageFile = 'grantB.gif'
   ,@ToolTip = 'Grant user permission'
go

execute dbo.shlActionRulesInsert
    @ObjectName = 'shlUserPerm'
   ,@ActionName = 'Grant'
   ,@RuleName = 'R1'
   ,@FieldName = 'Type'
   ,@Value = 'GA'
   ,@ValidationType = 'NE'
go

execute dbo.shlActionsInsert
    @ObjectName = 'shlUserPerm'
   ,@ActionName = 'Revoke'
   ,@Process = 'shlModuleUserRevoke'
   ,@RowBased = 'Y'
   ,@ImageFile = 'NoneAB.gif'
   ,@ToolTip = 'Revoke user permission'
go

execute dbo.shlActionRulesInsert
    @ObjectName = 'shlUserPerm'
   ,@ActionName = 'Revoke'
   ,@RuleName = 'R1'
   ,@FieldName = 'Type'
   ,@Value = 'NA'
   ,@ValidationType = 'NE'
go

execute dbo.shlActionsInsert
    @ObjectName = 'shlUserPerm'
   ,@ActionName = 'Deny'
   ,@Process = 'shlModuleUserDeny'
   ,@RowBased = 'Y'
   ,@ImageFile = 'denyB.gif'
   ,@ToolTip = 'Deny user permission'
go

execute dbo.shlActionRulesInsert
    @ObjectName = 'shlUserPerm'
   ,@ActionName = 'Deny'
   ,@RuleName = 'R1'
   ,@FieldName = 'Type'
   ,@Value = 'DA'
   ,@ValidationType = 'NE'
go

execute dbo.shlActionsInsert
    @ObjectName = 'shlUserPerm'
   ,@ActionName = 'Delete'
   ,@Process = 'shlModuleUserDelete'
   ,@RowBased = 'Y'
   ,@ImageFile = 'delete.gif'
   ,@ToolTip = 'Remove user permission'
go

execute dbo.shlActionRulesInsert
    @ObjectName = 'shlUserPerm'
   ,@ActionName = 'Delete'
   ,@RuleID = 0
   ,@RuleName = 'RD'
   ,@FieldName = 'Type'
   ,@Value = 'GA'
go

execute dbo.shlActionRulesInsert
    @ObjectName = 'shlUserPerm'
   ,@ActionName = 'Delete'
   ,@RuleID = 1
   ,@RuleName = 'RD'
   ,@FieldName = 'Type'
   ,@Value = 'DA'
go

execute dbo.shlActionRulesInsert
    @ObjectName = 'shlUserPerm'
   ,@ActionName = 'Delete'
   ,@RuleID = 2
   ,@RuleName = 'RD'
   ,@FieldName = 'Type'
   ,@Value = 'NA'
go

print '.oOo.'
go

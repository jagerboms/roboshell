print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlGridFormInsert
    @ObjectName = 'shlRoleMembers'
   ,@Title = 'Role Members'
   ,@DataParameter = 'shlRoleMembersGet'
   ,@TitleParameters = 'UserName'
   ,@HelpPage = 'RoleMembers.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'shlRoleMembers'
   ,@ModuleID = 'security'
   ,@ObjectName = 'shlRoleMembers'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'shlRoleMembers'
   ,@ParameterName = 'UserName'
   ,@ValueType = 'String'
   ,@Width = 128
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'shlRoleMembers'
   ,@FieldName = 'MemberName'
   ,@Label = 'Member'
   ,@Width = 128
   ,@DisplayWidth = 200
   ,@ValueType = 'String'
   ,@IsPrimary = 'Y'
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'shlRoleMembers'
   ,@ActionName = 'Refresh'
   ,@Process = 'shlRoleMembersGet'
   ,@ImageFile = 'Refresh.gif'
   ,@ToolTip = 'Refresh member details'
   ,@KeyCode = 116
go

print '.oOo.'
go

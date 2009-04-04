print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlGridFormInsert
    @ObjectName = 'shlGroupMember'
   ,@Title = 'Group Members'
   ,@DataParameter = 'data'
   ,@TitleParameters = 'UserName||GroupName'
   ,@HelpPage = 'GroupMembers.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'shlGroupMember'
   ,@ModuleID = 'security'
   ,@ObjectName = 'shlGroupMember'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'shlGroupMember'
   ,@ParameterName = 'UserName'
   ,@ValueType = 'String'
   ,@Width = 128
go

execute dbo.shlParametersInsert
    @ObjectName = 'shlGroupMember'
   ,@ParameterName = 'LoginName'
   ,@ValueType = 'String'
   ,@Width = 128
go

execute dbo.shlParametersInsert
    @ObjectName = 'shlGroupMember'
   ,@ParameterName = 'GroupName'
   ,@ValueType = 'String'
   ,@Width = 128
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'shlGroupMember'
   ,@FieldName = 'MemberName'
   ,@Label = 'Member'
   ,@Width = 128
   ,@DisplayWidth = 200
   ,@ValueType = 'String'
   ,@IsPrimary = 'Y'
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'shlGroupMember'
   ,@ActionName = 'Refresh'
   ,@Process = 'shlGroupMemberGet'
   ,@ImageFile = 'Refresh.gif'
   ,@ToolTip = 'Refresh details'
go

execute dbo.shlActionsInsert
    @ObjectName = 'shlGroupMember'
   ,@ActionName = 'Exit'
   ,@CloseObject = 'Y'
   ,@ImageFile = 'Cancel.gif'
   ,@ToolTip = 'Exit'
   ,@KeyCode = 27
go

print '.oOo.'
go

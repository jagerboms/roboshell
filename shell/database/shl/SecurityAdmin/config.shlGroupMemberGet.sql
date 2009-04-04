print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlCallAsmInsert
    @ObjectName = 'shlGroupMemberGet'
   ,@LibraryName = 'ActiveDir'
   ,@ClassName = 'Members'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'shlGroupMemberGet'
   ,@ModuleID = 'security'
   ,@ObjectName = 'shlGroupMemberGet'
   ,@UpdateParent = 'Y'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'shlGroupMemberGet'
   ,@ParameterName = 'LoginName'
   ,@ValueType = 'String'
   ,@Width = 128
go

execute dbo.shlParametersInsert
    @ObjectName = 'shlGroupMemberGet'
   ,@ParameterName = 'GroupName'
   ,@ValueType = 'String'
   ,@Width = 128
   ,@IsInput = 'N'
go

execute dbo.shlParametersInsert
    @ObjectName = 'shlGroupMemberGet'
   ,@ParameterName = 'data'
   ,@ValueType = 'Object'
   ,@IsInput = 'N'
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlGroupMemberGet'
   ,@PropertyType = 'cr'
   ,@PropertyName = 'LoginName'
   ,@Value = ''
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlGroupMemberGet'
   ,@PropertyType = 'op'
   ,@PropertyName = 'data'
   ,@Value = 'data'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlGroupMemberGet'
   ,@PropertyType = 'op'
   ,@PropertyName = 'GroupName'
   ,@Value = 'GroupName'
go

print '.oOo.'
go

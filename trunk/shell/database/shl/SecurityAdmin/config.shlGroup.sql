print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlCallASMInsert
    @ObjectName = 'shlGroups'
   ,@LibraryName = 'ActiveDir'
   ,@ClassName = 'Groups'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'shlGroups'
   ,@ModuleID = 'Public'
   ,@ObjectName = 'shlGroups'
   ,@UpdateParent = 'Y'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'shlGroups'
   ,@ParameterName = 'shlGroups'
   ,@ValueType = 'Object'
   ,@IsInput = 'N'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlGroups'
   ,@PropertyType = 'op'
   ,@PropertyName = 'shlGroups'
   ,@Value = 'data'
go

print '.oOo.'
go

print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlCallASMInsert
    @ObjectName = 'shlUsers'
   ,@LibraryName = 'ActiveDir'
   ,@ClassName = 'Users'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'shlUsers'
   ,@ModuleID = 'Public'
   ,@ObjectName = 'shlUsers'
   ,@UpdateParent = 'Y'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'shlUsers'
   ,@ParameterName = 'shlUsers'
   ,@ValueType = 'Object'
   ,@IsInput = 'N'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlUsers'
   ,@PropertyType = 'op'
   ,@PropertyName = 'shlUsers'
   ,@Value = 'data'
go

print '.oOo.'
go

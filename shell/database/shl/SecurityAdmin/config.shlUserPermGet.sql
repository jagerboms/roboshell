print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @ObjectName = 'shlUserPermGet'
   ,@ProcName = 'shlUserPermGet'
   ,@DataParameter = 'data'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'shlUserPermGet'
   ,@ModuleID = 'security'
   ,@ObjectName = 'shlUserPermGet'
   ,@UpdateParent = 'Y'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'shlUserPermGet'
   ,@ParameterName = 'UserName'
   ,@ValueType = 'String'
   ,@Width = 128
go

print '.oOo.'
go

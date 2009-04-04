print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @ObjectName = 'shlRoleMembersGet'
   ,@ProcName = 'shlRoleMembersGet'
   ,@DataParameter = 'shlRoleMembersGet'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'shlRoleMembersGet'
   ,@ModuleID = 'security'
   ,@ObjectName = 'shlRoleMembersGet'
   ,@UpdateParent = 'Y'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'shlRoleMembersGet'
   ,@ParameterName = 'UserName'
   ,@ValueType = 'String'
   ,@Width = 128
go

print '.oOo.'
go

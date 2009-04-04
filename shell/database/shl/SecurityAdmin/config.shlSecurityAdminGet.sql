print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @ObjectName = 'shlSecurityAdminGet'
   ,@ProcName = 'shlSecurityAdminGet'
   ,@DataParameter = 'shlSecurityAdmin'
go

execute dbo.shlProcessesInsert
    @ProcessName = 'shlSecurityAdminGet'
   ,@ModuleID = 'security'
   ,@ObjectName = 'shlSecurityAdminGet'
   ,@UpdateParent = 'Y'
go

print '.oOo.'
go

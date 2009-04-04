print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlModuleProceduresInsert
    @ProcedureName = 'shlShellGet'
   ,@ModuleID = 'public'
go

execute dbo.shlModuleProceduresInsert
    @ProcedureName = 'shlUserPropertyAlter'
   ,@ModuleID = 'public'
go

print '.oOo.'
go

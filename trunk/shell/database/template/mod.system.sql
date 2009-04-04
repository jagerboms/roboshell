print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlModulesInsert
    @ModuleID = 'system'
   ,@OwnerModule = 'base'
   ,@Description = 'system'
go

print '.oOo.'
go

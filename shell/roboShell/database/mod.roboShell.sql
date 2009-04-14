print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlModulesInsert
    @ModuleID = 'system'
   ,@OwnerModule = 'base'
   ,@Description = 'Robo Shell'
go

execute dbo.shlModulesInsert
    @ModuleID = 'statics'
   ,@OwnerModule = 'system'
   ,@Description = 'Statics'
go

print '.oOo.'
go

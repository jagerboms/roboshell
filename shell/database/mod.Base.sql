print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlModulesInsert
    @ModuleID = 'base'
   ,@OwnerModule = 'base'
   ,@Description = 'base'
go

execute dbo.shlModulesInsert
    @ModuleID = 'public'
   ,@OwnerModule = 'base'
   ,@Description = 'Public'
go

execute dbo.shlModulesInsert
    @ModuleID = 'security'
   ,@OwnerModule = 'base'
   ,@Description = 'Security'
go

execute dbo.shlModulesInsert
    @ModuleID = 'system'
   ,@OwnerModule = 'base'
   ,@Description = 'System'
go

print '.oOo.'
go

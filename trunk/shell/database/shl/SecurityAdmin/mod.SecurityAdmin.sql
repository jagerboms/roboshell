print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlModulesInsert
    @ModuleID = 'securityadmin'
   ,@OwnerModule = 'security'
   ,@Description = 'Administer'
go

print '.oOo.'
go

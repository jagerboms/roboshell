print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlModulesInsert
    @ModuleID = 'helpsystem'
   ,@OwnerModule = 'statics'
   ,@Description = 'Systems'
go

execute dbo.shlModulesInsert
    @ModuleID = 'helpsystemmaintain'
   ,@OwnerModule = 'helpsystem'
   ,@Description = 'Maintain'
go

print '.oOo.'
go

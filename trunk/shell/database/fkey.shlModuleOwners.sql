print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if object_id('shlModuleOwners_Module') is not null
begin
    alter table dbo.shlModuleOwners drop constraint shlModuleOwners_Module
end
go
if object_id('shlModuleOwners_Owner') is not null
begin
    alter table dbo.shlModuleOwners drop constraint shlModuleOwners_Owner
end

alter table dbo.shlModuleOwners add constraint shlModuleOwners_Module
foreign key (ModuleID) references dbo.shlModules (ModuleID)
go

alter table dbo.shlModuleOwners add constraint shlModuleOwners_Owner
foreign key (OwnerModule) references dbo.shlModules (ModuleID)
go

print '.oOo.'
go

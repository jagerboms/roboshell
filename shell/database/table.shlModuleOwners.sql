print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if object_id('dbo.shlModuleOwners') is null
begin
    create table dbo.shlModuleOwners
    (
        ModuleID     varchar(32) not null
       ,OwnerModule  varchar(32) not null
       ,constraint shlModuleOwnersPK primary key clustered
        (
            ModuleID
        )
    )
    print 'new table dbo.shlModuleOwners'
end
go

print '.oOo.'
go

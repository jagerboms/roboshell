print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if object_id('dbo.shlModules') is null
begin
    create table dbo.shlModules
    (
        ModuleID      varchar(32) not null
       ,Description   varchar(50) not null
       ,constraint shlModulesPK primary key clustered
        (
            ModuleID
        )
    )
    print 'new table dbo.shlModules'
end
go

print '.oOo.'
go

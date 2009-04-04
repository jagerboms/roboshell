print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlModuleUsers') is null
begin
    create table dbo.shlModuleUsers
    (
        ModuleID   varchar(32) not null
       ,UserName   sysname not null
       ,GrantDeny  char(1) not null
       ,constraint shlModuleUsersPK primary key clustered
        (
            ModuleID
           ,UserName
        )
    )
end
go

print '.oOo.'
go

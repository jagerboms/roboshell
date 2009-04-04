print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if object_id('dbo.shlModuleUsersAudit') is null
begin
    print 'creating dbo.shlModuleUsersAudit'
    create table dbo.shlModuleUsersAudit
    (
        ModuleID varchar(32) not null
       ,UserName sysname not null        -- sysusers.name
       ,AuditID  integer not null
       ,GrantDeny char(1) not null
       ,ActionType char(1) not null
       ,AuditTime datetime not null
       ,UserID sysname not null
       ,constraint shlModuleUsersAuditPK primary key clustered
        (
            ModuleID
           ,UserName
           ,AuditID
        )
    )
end
go

print '.oOo.'
go

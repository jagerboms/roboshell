print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

if object_id('dbo.helpObjectsAudit') is null
begin
    print 'creating dbo.helpObjectsAudit'
    create table dbo.helpObjectsAudit
    (
        SystemID varchar(12) not null
       ,ObjectName varchar(32) not null
       ,AuditID integer not null
       ,HelpText varchar(4000) null
       ,ColourText varchar(2000) null
       ,State varchar(2) not null
       ,ActionType char(1) not null
       ,AuditTime datetime not null constraint helpObjectsAuditAuditTime default (getdate())
       ,UserID sysname not null constraint helpObjectsAuditUserID default (suser_sname())
       ,constraint helpObjectsAuditPK primary key clustered
        (
            SystemID
           ,ObjectName
           ,AuditID
        )
    )
end
go

print '.oOo.'
go

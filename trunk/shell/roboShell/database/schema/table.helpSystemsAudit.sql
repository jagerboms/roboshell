print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

if object_id('dbo.helpSystemsAudit') is null
begin
    print 'creating dbo.helpSystemsAudit'
    create table dbo.helpSystemsAudit
    (
        SystemID varchar(12) not null
       ,AuditID integer not null
       ,SystemName varchar(100) null
       ,Copyright varchar(100) not null
       ,State varchar(2) not null
       ,ActionType char(1) not null
       ,AuditTime datetime not null constraint helpSystemsAuditAuditTime default (getdate())
       ,UserID sysname not null constraint helpSystemsAuditUserID default (suser_sname())
       ,constraint helpSystemsAuditPK primary key clustered
        (
            SystemID
           ,AuditID
        )
    )
end
go

print '.oOo.'
go

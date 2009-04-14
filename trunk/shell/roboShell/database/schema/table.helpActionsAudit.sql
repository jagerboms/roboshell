print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

if object_id('dbo.helpActionsAudit') is null
begin
    print 'creating dbo.helpActionsAudit'
    create table dbo.helpActionsAudit
    (
        SystemID varchar(12) not null
       ,ObjectName varchar(32) not null
       ,ActionName varchar(32) not null
       ,AuditID integer not null
       ,HelpText varchar(4000) null
       ,State varchar(2) not null
       ,ActionType char(1) not null
       ,AuditTime datetime not null constraint helpActionsAuditAuditTime default (getdate())
       ,UserID sysname not null constraint helpActionsAuditUserID default (suser_sname())
       ,constraint helpActionsAuditPK primary key clustered
        (
            SystemID
           ,ObjectName
           ,ActionName
           ,AuditID
        )
    )
end
go

print '.oOo.'
go

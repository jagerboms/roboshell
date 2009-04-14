print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

if object_id('dbo.helpFieldsAudit') is null
begin
    print 'creating dbo.helpFieldsAudit'
    create table dbo.helpFieldsAudit
    (
        SystemID varchar(12) not null
       ,ObjectName varchar(32) not null
       ,FieldName varchar(32) not null
       ,AuditID integer not null
       ,HelpText varchar(4000) null
       ,State varchar(2) not null
       ,ActionType char(1) not null
       ,AuditTime datetime not null constraint helpFieldsAuditAuditTime default (getdate())
       ,UserID sysname not null constraint helpFieldsAuditUserID default (suser_sname())
       ,constraint helpFieldsAuditPK primary key clustered
        (
            SystemID
           ,ObjectName
           ,FieldName
           ,AuditID
        )
    )
end
go

print '.oOo.'
go

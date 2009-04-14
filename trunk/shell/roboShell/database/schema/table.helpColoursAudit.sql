print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

if object_id('dbo.helpColoursAudit') is null
begin
    print 'creating dbo.helpColoursAudit'
    create table dbo.helpColoursAudit
    (
        SystemID varchar(12) not null
       ,ObjectName varchar(32) not null
       ,ColourValue varchar(200) not null
       ,AuditID integer not null
       ,ValueDescription varchar(30) null
       ,State varchar(2) not null
       ,ActionType char(1) not null
       ,AuditTime datetime not null constraint helpColoursAuditAuditTime default (getdate())
       ,UserID sysname not null constraint helpColoursAuditUserID default (suser_sname())
       ,constraint helpColoursAuditPK primary key clustered
        (
            SystemID
           ,ObjectName
           ,ColourValue
           ,AuditID
        )
    )
end
go

print '.oOo.'
go

print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if object_id('dbo.shlVariablesAudit') is null 
begin
    print 'creating dbo.shlVariablesAudit'
    create table dbo.shlVariablesAudit
    (
        VariableID varchar(32) not null
       ,AuditID integer not null
       ,VariableValue varchar(200) null
       ,ShellUse   char(1) not null
       ,ActionType char(1) not null
       ,AuditTime datetime not null
       ,UserID sysname not null
       ,constraint shlVariablesAuditPK primary key clustered
        (
            VariableID
           ,AuditID
        )
    )
end
go

print '.oOo.'
go

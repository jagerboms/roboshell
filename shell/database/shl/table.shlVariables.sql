print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if object_id('dbo.shlVariables') is null 
begin
    print 'creating dbo.shlVariables'
    create table dbo.shlVariables
    (
        VariableID   varchar(32)  not null
       ,VariableValue varchar(200) null
       ,ShellUse      char(1) not null
       ,AuditID      integer not null
       ,constraint shlVariablesPK primary key clustered
        (
            VariableID
        )
    )
end
go

print '.oOo.'
go

print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlActionProcessRules') is null
begin
    create table dbo.shlActionProcessRules
    (
        ObjectName varchar(32) not null
       ,ActionName varchar(32) not null
       ,Value varchar(200) not null
       ,Process varchar(32) null
       ,constraint shlActionProcessRulesPK primary key clustered
       (
            ObjectName
           ,ActionName
           ,Value
       )
    )
end
go

print '.oOo.'
go

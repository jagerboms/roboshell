print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlValidationRules') is null
begin
    create table dbo.shlValidationRules
    (
        ObjectName varchar(32) not null
       ,ValidationName varchar(32) not null
       ,FieldName varchar(32) not null
       ,constraint shlValidationRulesPK primary key clustered
       (
            ObjectName
           ,ValidationName
           ,FieldName
       )
    )
end
go

print '.oOo.'
go

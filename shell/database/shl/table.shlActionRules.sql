print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlActionRules') is null
begin
    create table dbo.shlActionRules
    (
        ObjectName varchar(32) not null
       ,ActionName varchar(32) not null
       ,RuleID integer not null
       ,RuleName varchar(32) not null
       ,FieldName varchar(32) not null
       ,ValidationType char(2) not null   -- EQ,NE,NN,GT,GE,LT,LE
       ,Value varchar(200) not null
       ,constraint shlActionRulesPK primary key clustered
       (
            ObjectName
           ,ActionName
           ,RuleID
       )
    )
end
go

print '.oOo.'
go

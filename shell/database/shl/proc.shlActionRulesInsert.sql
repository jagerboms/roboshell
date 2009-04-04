print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlActionRulesInsert') is not null
begin
    drop procedure dbo.shlActionRulesInsert
end
go

create procedure dbo.shlActionRulesInsert
    @ObjectName varchar(32)
   ,@ActionName varchar(32)
   ,@RuleID integer = 0
   ,@RuleName varchar(32)
   ,@FieldName varchar(32)
   ,@Value varchar(200)
   ,@ValidationType char(2) = 'EQ'
as
begin
    set nocount on
    declare @e integer

    set @e = 0
    while @e = 0
    begin
        print 'Rule: ' + rtrim(@ObjectName) + '.' + rtrim(@ActionName) + '.' + @RuleName

        insert into shlActionRules
        (
            ObjectName, ActionName, RuleID, RuleName,
            FieldName, ValidationType, Value
        )
        values
        (
            @ObjectName, @ActionName, @RuleID, @RuleName,
            @FieldName, @ValidationType, @Value
        )
        set @e = @@error
        break
    end
end
go

print '.oOo.'
go

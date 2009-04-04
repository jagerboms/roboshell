print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlValidationRulesInsert') is not null
begin
    drop procedure dbo.shlValidationRulesInsert
end
go

create procedure dbo.shlValidationRulesInsert
    @ObjectName varchar(32)
   ,@ValidationName varchar(32)
   ,@FieldName varchar(32)
as
begin
    set nocount on
    declare @e integer

    set @e = 0
    while @e = 0
    begin
        print 'Validation Rule: ' + rtrim(@ObjectName) + '.' + rtrim(@ValidationName)
                 + '.' + rtrim(@FieldName)

        insert into dbo.shlValidationRules
        (
            ObjectName, ValidationName, FieldName
        )
        values
        (
            @ObjectName, @ValidationName, @FieldName
        )
        set @e = @@error
        break
    end
end
go

print '.oOo.'
go

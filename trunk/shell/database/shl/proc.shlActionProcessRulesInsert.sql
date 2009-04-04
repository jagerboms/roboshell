print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if object_id('dbo.shlActionProcessRulesInsert') is not null
begin
    drop procedure dbo.shlActionProcessRulesInsert
end
go

create procedure dbo.shlActionProcessRulesInsert
    @ObjectName varchar(32)
   ,@ActionName varchar(32)
   ,@Value varchar(200)
   ,@Process varchar(32)
   ,@Update char(1) = 'N'
as
begin
    set nocount on
    declare @e integer
           ,@c integer

    set @e = 0
    while @e = 0
    begin
        print 'ProcessRule: ' + rtrim(@ObjectName) + '.' + rtrim(@ActionName) + '.' + @Value

        if upper(@Update) = 'Y'
        begin
            update  dbo.shlActionProcessRules
            set     Process = @Process
            where   ObjectName = @ObjectName
            and     ActionName = @ActionName
            and     Value = @Value

            select  @e = @@error
                   ,@c = @@rowcount
            if @e <> 0 or @c > 0
            begin
                break
            end
        end

        insert into dbo.shlActionProcessRules
        (
            ObjectName, ActionName, Value, Process
        )
        values
        (
            @ObjectName, @ActionName, @Value, @Process
        )
        set @e = @@error
        break
    end
    return @e
end
go

print '.oOo.'
go

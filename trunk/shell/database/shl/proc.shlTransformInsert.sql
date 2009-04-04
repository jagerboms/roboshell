print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlTransformInsert') is not null
begin
    drop procedure dbo.shlTransformInsert
end
go

create procedure dbo.shlTransformInsert
    @ObjectName varchar(32)
   ,@ChoiceParameter varchar(32) = null
   ,@NotFoundProcess varchar(32) = null
   ,@ProcessParameter varchar(32) = null
as
begin
    set nocount on
    declare @e integer

    set @e = 0
    while @e = 0
    begin
        begin transaction

        execute @e = dbo.shlObjectsInsert
            @ObjectName = @ObjectName
           ,@ObjectType = 'Transform'
        if @e <> 0
        begin
            break
        end

        if @ChoiceParameter is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'ChoiceParameter'
               ,@Value = @ChoiceParameter
            if @e <> 0
            begin
                break
            end
        end

        if @ProcessParameter is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'ProcessParameter'
               ,@Value = @ProcessParameter
            if @e <> 0
            begin
                break
            end
        end

        if @NotFoundProcess is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'NotFoundProcess'
               ,@Value = @NotFoundProcess
            if @e <> 0
            begin
                break
            end
        end
        break
    end
    if @e <> 0
    begin
        if @@trancount > 0
        begin
            rollback transaction
        end
    end
    else
    begin
        if @@trancount > 0
        begin
            commit transaction
        end
    end
    return @e
end
go

print '.oOo.'
go

print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlCallOutInsert') is not null
begin
    drop procedure dbo.shlCallOutInsert
end
go

create procedure dbo.shlCallOutInsert
    @ObjectName varchar(32)
   ,@ClassName varchar(128)
   ,@MethodName varchar(128)
   ,@ReturnParamName varchar(32) = null
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
           ,@ObjectType = 'CallOut'
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.shlPropertiesInsert
            @ObjectName = @ObjectName
           ,@PropertyName = 'ClassName'
           ,@Value = @ClassName
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.shlPropertiesInsert
            @ObjectName = @ObjectName
           ,@PropertyName = 'MethodName'
           ,@Value = @MethodName
        if @e <> 0
        begin
            break
        end

        if @ReturnParamName is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'ReturnParamName'
               ,@Value = @ReturnParamName
            if @e <> 0
            begin
                break
            end
        end

        set @e = @@error
        if @e <> 0
        begin
            break
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

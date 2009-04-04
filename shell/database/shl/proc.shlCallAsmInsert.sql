print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlCallAsmInsert') is not null
begin
    drop procedure dbo.shlCallAsmInsert
end
go

create procedure dbo.shlCallAsmInsert
    @ObjectName varchar( 32)
   ,@LibraryName varchar(128) = null
   ,@ClassName varchar(128) = null
   ,@MethodName varchar(128) = null
   ,@ObjectParamName varchar(32) = null
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
           ,@ObjectType = 'CallAsm'
        if @e <> 0
        begin
            break
        end

        if @LibraryName is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'LibraryName'
               ,@Value = @LibraryName
            if @e <> 0
            begin
                break
            end
        end

        if @ClassName is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'ClassName'
               ,@Value = @ClassName
            if @e <> 0
            begin
                break
            end
        end

        if @MethodName is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'MethodName'
               ,@Value = @MethodName
            if @e <> 0
            begin
                break
            end
        end

        if @ObjectParamName is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'ObjectParamName'
               ,@Value = @ObjectParamName
            if @e <> 0
            begin
                break
            end
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

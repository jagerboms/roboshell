print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if object_id('dbo.shlStoredProcInsert') is not null
begin
    drop procedure dbo.shlStoredProcInsert
end
go

create procedure dbo.shlStoredProcInsert
    @objectname varchar(32)
   ,@procname varchar(32)
   ,@connectkey varchar(32) = 'Default'
   ,@dataparameter varchar(32) = null
   ,@mode char(1) = 'D'  -- Data/eXecute only/output to Parameters
   ,@messages char(1) = null
   ,@timeout integer = null
as
begin
    set nocount on
    declare @e integer
           ,@m varchar(200)

    set @e = 0
    while @e = 0
    begin
        set @mode = upper( coalesce( @mode,'D' ) )
        if @mode not in ('X', 'P')
        begin
            set @mode = 'D'
        end

        if object_id(@procname) is null
        begin
            set @e = 50040
            set @m = 'Error: procedure ' + coalesce(@procname, 'null') 
                            + ' does not exist in the database'
            raiserror @e @m
            break
        end

        begin transaction

        execute @e = dbo.shlObjectsInsert
            @ObjectName = @objectname
           ,@ObjectType = 'StoredProc'
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.shlPropertiesInsert
            @ObjectName = @objectname
           ,@PropertyName = 'procname'
           ,@Value = @procname
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.shlPropertiesInsert
            @ObjectName = @objectname
           ,@PropertyName = 'connectkey'
           ,@Value = @connectkey
        if @e <> 0
        begin
            break
        end

        if @dataparameter is not null
        begin
            if @mode = 'X'
            begin
                set @e = 50041
                set @m = 'Error: Execute only object cannot have a data parameter defined.'
                raiserror @e @m
                break
            end

            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @objectname
               ,@PropertyName = 'dataparameter'
               ,@Value = @dataparameter
            if @e <> 0
            begin
                break
            end

            execute @e = dbo.shlParametersInsert
                @ObjectName = @objectname
               ,@ParameterName = @dataparameter
               ,@ValueType = 'Object'
               ,@IsInput = 'N'
            if @e <> 0
            begin
                break
            end
        end
        else
        begin
            if @Mode = 'D'
            begin
                set @e = 50042
                set @m = 'Error: Data parameter is not defined.'
                raiserror @e @m
                break
            end
        end

        execute @e = dbo.shlPropertiesInsert
            @ObjectName = @objectname
           ,@PropertyName = 'mode'
           ,@Value = @mode

        if @messages is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @objectname
               ,@PropertyName = 'messages'
               ,@Value = @messages
            if @e <> 0
            begin
                break
            end
        end

        if @timeout is not null
        begin
            set @m = cast(@timeout as varchar)
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @objectname
               ,@PropertyName = 'timeout'
               ,@Value = @m
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

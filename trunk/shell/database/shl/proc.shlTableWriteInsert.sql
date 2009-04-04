print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlTableWriteInsert') is not null
begin
    drop procedure dbo.shlTableWriteInsert
end
go

create procedure dbo.shlTableWriteInsert
    @ObjectName varchar(32)
   ,@DataParameter varchar(32) = 'data'
   ,@TableWritePreProcess varchar(32) = null
   ,@RowWriteProcess varchar(32)
   ,@TableWritePostProcess varchar(32) = null
   ,@ErrorProcess varchar(32) = null
as
begin
    set nocount on
    declare @e integer
           ,@m varchar(200)

    set @e = 0
    while @e = 0
    begin
        if @RowWriteProcess is null
        begin
            set @e = 60600
            set @m = 'Error row write process must be provided...'
            raiserror @e @m
            break
        end
        else
        begin
            if not exists
            (
                select  'a'
                from    dbo.shlProcesses p
                where   p.ProcessName = @RowWriteProcess
            )
            begin
                set @e = 60600
                set @m = 'Error row write process ' + @RowWriteProcess + ' does not exist...'
                raiserror @e @m
                break
            end
        end

        if @TableWritePreProcess is not null
        begin
            if not exists
            (
                select  'a'
                from    dbo.shlProcesses p
                where   p.ProcessName = @TableWritePreProcess
            )
            begin
                set @e = 60600
                set @m = 'Error table write pre-process ' + @TableWritePreProcess + ' does not exist...'
                raiserror @e @m
                break
            end
        end

        if @TableWritePostProcess is not null
        begin
            if not exists
            (
                select  'a'
                from    dbo.shlProcesses p
                where   p.ProcessName = @TableWritePostProcess
            )
            begin
                set @e = 60600
                set @m = 'Error table write post-process ' + @TableWritePostProcess + ' does not exist...'
                raiserror @e @m
                break
            end
        end

        begin transaction

        execute @e = dbo.shlObjectsInsert
            @ObjectName = @ObjectName
           ,@ObjectType = 'TableWrite'
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.shlPropertiesInsert
            @ObjectName = @ObjectName
           ,@PropertyName = 'DataParameter'
           ,@Value = @DataParameter
        if @e <> 0
        begin
            break
        end

        if @TableWritePreProcess is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'TableWritePreProcess'
               ,@Value = @TableWritePreProcess
            if @e <> 0
            begin
                break
            end
        end

        execute @e = dbo.shlPropertiesInsert
            @ObjectName = @ObjectName
           ,@PropertyName = 'RowWriteProcess'
           ,@Value = @RowWriteProcess
        if @e <> 0
        begin
            break
        end

        if @TableWritePostProcess is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'TableWritePostProcess'
               ,@Value = @TableWritePostProcess
            if @e <> 0
            begin
                break
            end
        end

        if @ErrorProcess is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'ErrorProcess'
               ,@Value = @ErrorProcess
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

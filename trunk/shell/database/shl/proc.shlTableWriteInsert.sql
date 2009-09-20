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
   ,@PreWriteProcess varchar(32) = null
   ,@WriteProcess varchar(32)
   ,@PostWriteProcess varchar(32) = null
as
begin
    set nocount on
    declare @e integer
           ,@m varchar(200)

    set @e = 0
    while @e = 0
    begin
        if @WriteProcess is null
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
                where   p.ProcessName = @WriteProcess
            )
            begin
                set @e = 60600
                set @m = 'Error row write process ' + @WriteProcess + ' does not exist...'
                raiserror @e @m
                break
            end
        end

        if @PreWriteProcess is not null
        begin
            if not exists
            (
                select  'a'
                from    dbo.shlProcesses p
                where   p.ProcessName = @PreWriteProcess
            )
            begin
                set @e = 60600
                set @m = 'Error write pre-process ' + @PreWriteProcess + ' does not exist...'
                raiserror @e @m
                break
            end
        end

        if @PostWriteProcess is not null
        begin
            if not exists
            (
                select  'a'
                from    dbo.shlProcesses p
                where   p.ProcessName = @PostWriteProcess
            )
            begin
                set @e = 60600
                set @m = 'Error table write post-process ' + @PostWriteProcess + ' does not exist...'
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

        if @PreWriteProcess is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'PreWriteProcess'
               ,@Value = @PreWriteProcess
            if @e <> 0
            begin
                break
            end
        end

        execute @e = dbo.shlPropertiesInsert
            @ObjectName = @ObjectName
           ,@PropertyName = 'WriteProcess'
           ,@Value = @WriteProcess
        if @e <> 0
        begin
            break
        end

        if @PostWriteProcess is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'PostWriteProcess'
               ,@Value = @PostWriteProcess
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

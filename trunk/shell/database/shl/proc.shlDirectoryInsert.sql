print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlDirectoryInsert') is not null
begin
    drop procedure dbo.shlDirectoryInsert
end
go

create procedure dbo.shlDirectoryInsert
    @ObjectName varchar(32)
   ,@Title varchar(100)
   ,@TitleParameters varchar(132) = null
   ,@InitialDirectory varchar(132) = null
   ,@OutputParameter varchar(32)
   ,@AllowNew char(1) = null
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
           ,@ObjectType = 'Directory'
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.shlPropertiesInsert
            @ObjectName = @ObjectName
           ,@PropertyName = 'Title'
           ,@Value = @Title
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.shlPropertiesInsert
            @ObjectName = @ObjectName
           ,@PropertyName = 'OutputParameter'
           ,@Value = @OutputParameter
        if @e <> 0
        begin
            break
        end

        if @TitleParameters is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'TitleParameters'
               ,@Value = @TitleParameters
            if @e <> 0
            begin
                break
            end
        end

        if @InitialDirectory is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'InitialDirectory'
               ,@Value = @InitialDirectory
            if @e <> 0
            begin
                break
            end
        end

        if coalesce(@AllowNew, 'Y') = 'N'
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'AllowNew'
               ,@Value = 'N'
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

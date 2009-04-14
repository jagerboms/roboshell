print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlFileOpenInsert') is not null
begin
    drop procedure dbo.shlFileOpenInsert
end
go

create procedure dbo.shlFileOpenInsert
    @ObjectName varchar(32)
   ,@Title varchar(100)
   ,@TitleParameters varchar(132) = null
   ,@InitialDirectory varchar(132) = null
   ,@Filter varchar(132) = null
   ,@FilterIndex integer = null
   ,@Multiselect char(1) = null     -- Y/N
   ,@OutputParameter varchar(32)
as
begin
    set nocount on
    declare @e integer
           ,@s varchar(100)

    set @e = 0
    while @e = 0
    begin
        begin transaction

        execute @e = dbo.shlObjectsInsert
            @ObjectName = @ObjectName
           ,@ObjectType = 'FileOpen'
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

        if @Filter is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'Filter'
               ,@Value = @Filter
            if @e <> 0
            begin
                break
            end
        end

        if coalesce(@FilterIndex, 0) > 0
        begin
            set @s = cast(@FilterIndex as varchar)
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'FilterIndex'
               ,@Value = @s
            if @e <> 0
            begin
                break
            end
        end

        if coalesce(@Multiselect, 'N') = 'Y'
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'Multiselect'
               ,@Value = 'Y'
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

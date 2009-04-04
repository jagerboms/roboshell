print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlDialogFormInsert') is not null
begin
    drop procedure dbo.shlDialogFormInsert
end
go

create procedure dbo.shlDialogFormInsert
    @ObjectName varchar(32)
   ,@Title varchar(100)
   ,@TitleParameters varchar(132) = null
   ,@HelpPage varchar(200) = null
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
           ,@ObjectType = 'Dialog'
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

        if @HelpPage is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'HelpPage'
               ,@Value = @HelpPage
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

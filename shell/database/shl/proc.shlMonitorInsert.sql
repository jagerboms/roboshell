print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlMonitorInsert') is not null
begin
    drop procedure dbo.shlMonitorInsert
end
go

create procedure dbo.shlMonitorInsert
    @ObjectName varchar(32)
   ,@Title varchar(200)
   ,@ServerParameter char(32)
   ,@ServiceParameter char(32)
   ,@TitleParameters varchar(132) = null
   ,@HelpPage varchar(200) = null
as
begin
    set nocount on
    declare @e integer

    set @e = 0
    while @e = 0
    begin
     
        if coalesce(@Title, '') = ''
        begin
            set @e = 50040
            raiserror @e 'Invalid Title'
            break
        end

        if coalesce(@ServerParameter, '') = ''
        begin
            set @e = 50041
            raiserror @e 'Invalid Server parameter'
            break
        end

        if coalesce(@ServiceParameter, '') = ''
        begin
            set @e = 50040
            raiserror @e 'Invalid Service parameter'
            break
        end

        begin transaction

        execute @e = dbo.shlObjectsInsert
            @ObjectName = @ObjectName
           ,@ObjectType = 'Monitor'
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
           ,@PropertyName = 'ServerParameter'
           ,@Value = @ServerParameter
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.shlPropertiesInsert
            @ObjectName = @ObjectName
           ,@PropertyName = 'ServiceParameter'
           ,@Value = @ServiceParameter
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

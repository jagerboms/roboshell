print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
go

if object_id('dbo.shlReportInsert') is not null
begin
    drop procedure dbo.shlReportInsert
end
go

create procedure dbo.shlReportInsert
    @ObjectName char(32)
   ,@Title varchar(200)
   ,@DataParameter char(32) = null
   ,@PrintPreview char(1) = 'N'
   ,@DefaultPrinter char(1) = 'N'
as
begin
    set nocount on
    declare @err integer

    set @err = 0
    while @err = 0
    begin
        set @PrintPreview = upper(@PrintPreview)
        if @PrintPreview <> 'Y'
        begin
            set @PrintPreview = 'N'
        end

        set @DefaultPrinter = upper(@DefaultPrinter)
        if @DefaultPrinter <> 'Y'
        begin
            set @DefaultPrinter = 'N'
        end

        if @Title is null
        begin
            set @err = 50040
            raiserror @err 'Invalid Title'
            break
        end

        begin transaction

        execute @err = shlObjectsInsert
            @ObjectName = @ObjectName
           ,@ObjectType = 'Report'
        if @err <> 0
        begin
            break
        end

        execute @err = shlPropertiesInsert
            @ObjectName = @ObjectName
           ,@PropertyName = 'Title'
           ,@Value = @Title
        if @err <> 0
        begin
            break
        end

        if @DataParameter is not null
        begin
            execute @err = shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'DataParameter'
               ,@Value = @DataParameter
            if @err <> 0
            begin
                break
            end
        end
       
        execute @err = shlPropertiesInsert
            @ObjectName = @ObjectName
           ,@PropertyName = 'PrintPreview'
           ,@Value = @PrintPreview
        if @err <> 0
        begin
            break
        end

        execute @err = shlPropertiesInsert
            @ObjectName = @ObjectName
           ,@PropertyName = 'DefaultPrinter'
           ,@Value = @DefaultPrinter
        if @err <> 0
        begin
            break
        end

        break
    end
    if @err <> 0
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
    return @err
end
go

print 'Complete...'
go

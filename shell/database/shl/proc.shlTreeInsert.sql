print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if object_id('dbo.shlTreeInsert') is not null
begin
    drop procedure dbo.shlTreeInsert
end
go

create procedure dbo.shlTreeInsert
    @ObjectName varchar(32)
   ,@Title varchar(100)                -- caption on tree form
   ,@DataParameter varchar(32) = 'data'  -- parameter with dataset (dataTable)
   ,@KeyColumn varchar(32)               -- key field in dataset
   ,@DescriptionColumn varchar(32)       -- dataset field to display
   ,@ParentColumn varchar(32)            -- field representing parent's key
   ,@TypeColumn varchar(32)              -- field used for mapping images
   ,@ColourColumn varchar(32) = null     -- field used for colour mapping
   ,@DefaultImage varchar(200) = null    -- path to image used as default
   ,@TitleParameters varchar(132) = null -- '||' delimited field list appended to the form caption
   ,@RefreshTree char(1) = 'N'           -- refreshes entire tree on listen notify
   ,@HelpPage varchar(200) = null        -- path to help file for this form
as
begin
    set nocount on
    declare @e integer
           ,@m varchar(200)

    set @e = 0
    while @e = 0
    begin

        begin transaction

        execute @e = dbo.shlObjectsInsert
            @ObjectName = @ObjectName
           ,@ObjectType = 'Tree'
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


        if @DataParameter is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'DataParameter'
               ,@Value = @DataParameter
            if @e <> 0
            begin
                break
            end

            execute @e = dbo.shlParametersInsert
                @ObjectName = @ObjectName
               ,@ParameterName = @DataParameter
               ,@ValueType = 'Object'
            if @e <> 0
            begin
                break
            end
        end
        else
        begin
            set @e = 50042
            set @m = 'Error: Data parameter is not defined.'
            raiserror @e @m
            break
        end

        execute @e = dbo.shlPropertiesInsert
            @ObjectName = @ObjectName
           ,@PropertyName = 'KeyColumn'
           ,@Value = @KeyColumn
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.shlPropertiesInsert
            @ObjectName = @ObjectName
           ,@PropertyName = 'DescriptionColumn'
           ,@Value = @DescriptionColumn
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.shlPropertiesInsert
            @ObjectName = @ObjectName
           ,@PropertyName = 'TypeColumn'
           ,@Value = @TypeColumn
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.shlPropertiesInsert
            @ObjectName = @ObjectName
           ,@PropertyName = 'ParentColumn'
           ,@Value = @ParentColumn
        if @e <> 0
        begin
            break
        end

        if @ColourColumn is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'ColourColumn'
               ,@Value = @ColourColumn
            if @e <> 0
            begin
                break
            end
        end

        if @DefaultImage is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'DefaultImage'
               ,@Value = @DefaultImage
            if @e <> 0
            begin
                break
            end
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

        if upper(coalesce(@RefreshTree, 'N')) = 'Y'
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'RefreshTree'
               ,@Value = 'Y'
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

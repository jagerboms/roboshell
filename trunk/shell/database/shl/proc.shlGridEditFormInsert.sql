print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlGridEditFormInsert') is not null
begin
    drop procedure dbo.shlGridEditFormInsert
end
go

create procedure dbo.shlGridEditFormInsert
    @ObjectName varchar(32)
   ,@Title varchar(100)
   ,@DataParameter varchar(32) = 'data'
   ,@DisplayParameter varchar(32) = 'display'
   ,@Transpose char(1) = 'N'
   ,@TitleParameters varchar(132) = null
   ,@AddRowAction varchar(32) = null
   ,@DeleteRowAction varchar(32) = null
as
begin
    set nocount on
    declare @e integer

    set @e = 0
    while @e = 0
    begin
        set @Transpose = upper(@Transpose)
        if @Transpose <> 'Y'
        begin
            set @Transpose = 'N'
        end

        begin transaction

        execute @e = dbo.shlObjectsInsert
            @ObjectName = @ObjectName
           ,@ObjectType = 'GridEdit'
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
           ,@PropertyName = 'DataParameter'
           ,@Value = @DataParameter
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.shlPropertiesInsert
            @ObjectName = @ObjectName
           ,@PropertyName = 'DisplayParameter'
           ,@Value = @DisplayParameter
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

        execute @e = dbo.shlPropertiesInsert
            @ObjectName = @ObjectName
           ,@PropertyName = 'Transpose'
           ,@Value = @Transpose
        if @e <> 0
        begin
            break
        end

        if @AddRowAction is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'AddRowAction'
               ,@Value = @AddRowAction
            if @e <> 0
            begin
                break
            end
        end

        if @DeleteRowAction is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'DeleteRowAction'
               ,@Value = @DeleteRowAction
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

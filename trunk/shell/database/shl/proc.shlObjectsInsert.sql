print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlObjectsInsert') is not null
begin
    drop procedure dbo.shlObjectsInsert
end
go

create procedure dbo.shlObjectsInsert
    @ObjectName varchar(32)
   ,@ObjectType varchar(32)
as
begin
    set nocount on
    declare @e integer
           ,@count integer
           ,@tran integer

    set @e = 0
    while @e = 0
    begin
        print rtrim(@ObjectType) + ': ' + @ObjectName

        set @tran = @@trancount

        if @tran = 0
        begin
            begin transaction
        end

        delete
        from    dbo.shlActionRules
        where   ObjectName = @ObjectName

        set @e = @@error
        if @e <> 0
        begin
            break
        end

        delete
        from    dbo.shlActionProcessRules
        where   ObjectName = @ObjectName

        set @e = @@error
        if @e <> 0
        begin
            break
        end

        delete
        from    dbo.shlActions
        where   ObjectName = @ObjectName

        set @e = @@error
        if @e <> 0
        begin
            break
        end

        delete
        from    dbo.shlValidationRules
        where   ObjectName = @ObjectName

        set @e = @@error
        if @e <> 0
        begin
            break
        end

        delete
        from    dbo.shlValidations
        where   ObjectName = @ObjectName

        set @e = @@error
        if @e <> 0
        begin
            break
        end

        delete
        from    dbo.shlProperties
        where   ObjectName = @ObjectName

        set @e = @@error
        if @e <> 0
        begin
            break
        end

        delete
        from    dbo.shlParameters
        where   ObjectName = @ObjectName

        set @e = @@error
        if @e <> 0
        begin
            break
        end

        delete
        from    dbo.shlFields
        where   ObjectName = @ObjectName

        set @e = @@error
        if @e <> 0
        begin
            break
        end

        update  dbo.shlObjects
        set     ObjectType = @ObjectType
        where   ObjectName = @ObjectName

        select  @e = @@error
               ,@count = @@rowcount
        if @e <> 0
        begin
            break
        end

        if @count = 0
        begin
            insert into dbo.shlObjects
            (
                ObjectName, ObjectType
            )
            values
            (
                @ObjectName, @ObjectType
            )
            set @e = @@error
            if @e <> 0
            begin
                break
            end
        end
        break
    end
    if @e <> 0
    begin
        if @@trancount > 0 and @tran = 0
        begin
            rollback transaction
        end
    end
    else
    begin
        if @@trancount > 0 and @tran = 0
        begin
            commit transaction
        end
    end
    return @e
end
go

print '.oOo.'
go

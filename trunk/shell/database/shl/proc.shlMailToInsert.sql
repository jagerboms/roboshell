print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlMailToInsert') is not null
begin
    drop procedure dbo.shlMailToInsert
end
go

create procedure dbo.shlMailToInsert
    @ObjectName varchar(32)
   ,@email varchar(200) = null
   ,@cc varchar(200) = null
   ,@bcc varchar(200) = null
   ,@subject varchar(200) = null
   ,@body varchar(200) = null
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
           ,@ObjectType = 'MailTo'
        if @e <> 0
        begin
            break
        end

        if @email is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'email'
               ,@Value = @email
            if @e <> 0
            begin
                break
            end
        end

        if @cc is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'cc'
               ,@Value = @cc
            if @e <> 0
            begin
                break
            end
        end

        if @bcc is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'bcc'
               ,@Value = @bcc
            if @e <> 0
            begin
                break
            end
        end

        if @subject is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'subject'
               ,@Value = @subject
            if @e <> 0
            begin
                break
            end
        end

        if @body is not null
        begin
            execute @e = dbo.shlPropertiesInsert
                @ObjectName = @ObjectName
               ,@PropertyName = 'body'
               ,@Value = @body
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

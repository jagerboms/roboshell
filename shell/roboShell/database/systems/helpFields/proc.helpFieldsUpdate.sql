print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpFieldsUpdate') is not null
begin
    drop procedure dbo.helpFieldsUpdate
end
go
create procedure dbo.helpFieldsUpdate
    @SystemID varchar(12)
   ,@ObjectName varchar(32)
   ,@FieldName varchar(32)
   ,@HelpText varchar(4000) = null
   ,@AuditID integer
as
begin
    set nocount on
    declare @e integer
           ,@AudID integer
           ,@OldHelpText varchar(4000)

    set @e = 0
    while @e = 0
    begin
        select  @OldHelpText = a.HelpText
               ,@AudID = a.AuditID
        from    dbo.helpFields a (holdlock)
        where   a.SystemID = @SystemID
        and     a.ObjectName = @ObjectName
        and     a.FieldName = @FieldName

        if @@rowcount = 0  -- not found
        begin
            set @e = 51001
            raiserror (@e, 16, 1, 'Field')
            break
        end

        if @AudID <> @AuditID   -- already changed
        begin
            set @AudID = -1
            break
        end
        set @AudID = @AudID + 1

        if coalesce(@OldHelpText, '') = coalesce(@HelpText, '')
        begin
            set @e = 51002
            raiserror (@e, 16, 1)
            break
        end

        begin transaction

        update  dbo.helpFields
        set     HelpText = @HelpText
               ,AuditID = @AudID
        where   SystemID = @SystemID
        and     ObjectName = @ObjectName
        and     FieldName = @FieldName

        set @e = @@error
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.helpFieldsAuditInsert
            @SystemID = @SystemID
           ,@ObjectName = @ObjectName
           ,@FieldName = @FieldName
           ,@AuditID = @AudID
           ,@ActionType = 'U'
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
        execute dbo.helpFieldsGet    -- return the changes
            @pSystemID = @SystemID
           ,@pObjectName = @ObjectName
           ,@pFieldName = @FieldName

        if @AudID = -1
        begin
            print 'This Field has changed since it was retrieved'
        end
    end
    return @e
end
go
go

print '.oOo.'
go

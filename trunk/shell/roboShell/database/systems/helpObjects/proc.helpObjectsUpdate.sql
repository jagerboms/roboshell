print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpObjectsUpdate') is not null
begin
    drop procedure dbo.helpObjectsUpdate
end
go
create procedure dbo.helpObjectsUpdate
    @SystemID varchar(12)
   ,@ObjectName varchar(32)
   ,@HelpText varchar(4000) = null
   ,@ColourText varchar(2000) = null
   ,@AuditID integer
as
begin
    set nocount on
    declare @e integer
           ,@AudID integer
           ,@OldHelpText varchar(4000)
           ,@OldColourText varchar(2000)

    set @e = 0
    while @e = 0
    begin
        select  @OldHelpText = a.HelpText
               ,@OldColourText = a.ColourText
               ,@AudID = a.AuditID
        from    dbo.helpObjects a (holdlock)
        where   a.SystemID = @SystemID
        and     a.ObjectName = @ObjectName

        if @@rowcount = 0  -- not found
        begin
            set @e = 51001
            raiserror (@e, 16, 1, 'Object')
            break
        end

        if @AudID <> @AuditID   -- already changed
        begin
            set @AudID = -1
            break
        end
        set @AudID = @AudID + 1

        if coalesce(@OldHelpText, '') = coalesce(@HelpText, '')
        and coalesce(@OldColourText, '') = coalesce(@ColourText, '')
        begin
            set @e = 51002
            raiserror (@e, 16, 1)
            break
        end

        begin transaction

        update  dbo.helpObjects
        set     HelpText = @HelpText
               ,ColourText = @ColourText
               ,AuditID = @AudID
        where   SystemID = @SystemID
        and     ObjectName = @ObjectName

        set @e = @@error
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.helpObjectsAuditInsert
            @SystemID = @SystemID
           ,@ObjectName = @ObjectName
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
        execute dbo.helpObjectsGet    -- return the changes
            @pSystemID = @SystemID
           ,@pObjectName = @ObjectName

        if @AudID = -1
        begin
            print 'This Object has changed since it was retrieved'
        end
    end
    return @e
end
go
go

print '.oOo.'
go

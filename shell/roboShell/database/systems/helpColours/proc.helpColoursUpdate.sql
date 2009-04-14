print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpColoursUpdate') is not null
begin
    drop procedure dbo.helpColoursUpdate
end
go
create procedure dbo.helpColoursUpdate
    @SystemID varchar(12)
   ,@ObjectName varchar(32)
   ,@ColourValue varchar(200)
   ,@ValueDescription varchar(30) = null
   ,@AuditID integer
as
begin
    set nocount on
    declare @e integer
           ,@AudID integer
           ,@OldValueDescription varchar(30)

    set @e = 0
    while @e = 0
    begin
        select  @OldValueDescription = a.ValueDescription
               ,@AudID = a.AuditID
        from    dbo.helpColours a (holdlock)
        where   a.SystemID = @SystemID
        and     a.ObjectName = @ObjectName
        and     a.ColourValue = @ColourValue

        if @@rowcount = 0  -- not found
        begin
            set @e = 51001
            raiserror (@e, 16, 1, 'Colour')
            break
        end

        if @AudID <> @AuditID   -- already changed
        begin
            set @AudID = -1
            break
        end
        set @AudID = @AudID + 1

        if coalesce(@OldValueDescription, '') = coalesce(@ValueDescription, '')
        begin
            set @e = 51002
            raiserror (@e, 16, 1)
            break
        end

        begin transaction

        update  dbo.helpColours
        set     ValueDescription = @ValueDescription
               ,AuditID = @AudID
        where   SystemID = @SystemID
        and     ObjectName = @ObjectName
        and     ColourValue = @ColourValue

        set @e = @@error
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.helpColoursAuditInsert
            @SystemID = @SystemID
           ,@ObjectName = @ObjectName
           ,@ColourValue = @ColourValue
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
        execute dbo.helpColoursGet    -- return the changes
            @pSystemID = @SystemID
           ,@pObjectName = @ObjectName
           ,@pColourValue = @ColourValue

        if @AudID = -1
        begin
            print 'This Colour has changed since it was retrieved'
        end
    end
    return @e
end
go
go

print '.oOo.'
go

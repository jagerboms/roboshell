print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpSystemsUpdate') is not null
begin
    drop procedure dbo.helpSystemsUpdate
end
go
create procedure dbo.helpSystemsUpdate
    @SystemID varchar(12)
   ,@SystemName varchar(100) = null
   ,@Copyright varchar(100)
   ,@AuditID integer
as
begin
    set nocount on
    declare @e integer
           ,@AudID integer
           ,@OldSystemName varchar(100)
           ,@OldCopyright varchar(100)

    set @e = 0
    while @e = 0
    begin
        select  @OldSystemName = a.SystemName
               ,@OldCopyright = a.Copyright
               ,@AudID = a.AuditID
        from    dbo.helpSystems a (holdlock)
        where   a.SystemID = @SystemID

        if @@rowcount = 0  -- not found
        begin
            set @e = 51001
            raiserror (@e, 16, 1, 'System')
            break
        end

        if @AudID <> @AuditID   -- already changed
        begin
            set @AudID = -1
            break
        end
        set @AudID = @AudID + 1

        if coalesce(@OldSystemName, '') = coalesce(@SystemName, '')
        and coalesce(@OldCopyright, '') = coalesce(@Copyright, '')
        begin
            set @e = 51002
            raiserror (@e, 16, 1)
            break
        end

        begin transaction

        update  dbo.helpSystems
        set     SystemName = @SystemName
               ,Copyright = @Copyright
               ,AuditID = @AudID
        where   SystemID = @SystemID

        set @e = @@error
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.helpSystemsAuditInsert
            @SystemID = @SystemID
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
        execute dbo.helpSystemsGet    -- return the changes
            @pSystemID = @SystemID

        if @AudID = -1
        begin
            print 'This System has changed since it was retrieved'
        end
    end
    return @e
end
go
go

print '.oOo.'
go

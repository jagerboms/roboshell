print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpSystemsDisable') is not null
begin
    drop procedure dbo.helpSystemsDisable
end
go
create procedure dbo.helpSystemsDisable
    @SystemID varchar(12)
   ,@AuditID integer
as
begin
    set nocount on
    declare @e integer
           ,@AudID integer

    set @e = 0
    while @e = 0
    begin
        select  @AudID = a.AuditID
        from    dbo.helpSystems a (holdlock)
        where   a.SystemID = @SystemID

        if @@rowcount = 0
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

        begin transaction

        update  dbo.helpSystems
        set     State = 'dl'
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
           ,@ActionType = 'D'
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

print '.oOo.'
go

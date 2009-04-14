print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpSystemsInsert') is not null
begin
    drop procedure dbo.helpSystemsInsert
end
go
create procedure dbo.helpSystemsInsert
    @SystemID varchar(12)
   ,@SystemName varchar(100) = null
   ,@Copyright varchar(100)
as
begin
    set nocount on
    declare @e integer
           ,@AuditID integer
           ,@State char(2)

    set @e = 0
    while @e = 0
    begin
        set @SystemID = upper(@SystemID)

        begin transaction

        select  @AuditID = a.AuditID
               ,@State = a.State
        from    dbo.helpSystems a
        where   a.SystemID = @SystemID

        if @@rowcount > 0
        begin
            set @AuditID = @AuditID + 1

            if @State = 'ac'
            begin
                set @e = 51000
                raiserror (@e, 16, 1, 'System')
                break
            end

            update  dbo.helpSystems
            set     SystemName = @SystemName
                   ,Copyright = @Copyright
                   ,State = 'ac'
                   ,AuditID = @AuditID
            where   SystemID = @SystemID

            set @e = @@error
        end
        else
        begin
            set @AuditID = 1

            insert into dbo.helpSystems
            (
                SystemID, SystemName, Copyright,
                State, AuditID
            )
            values
            (
                @SystemID, @SystemName, @Copyright,
                'ac', @AuditID
            )
            set @e = @@error
        end
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.helpSystemsAuditInsert
            @SystemID = @SystemID
           ,@AuditID = @AuditID
           ,@ActionType = 'I'
        if @e <> 0
        begin
            break
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
go

print '.oOo.'
go

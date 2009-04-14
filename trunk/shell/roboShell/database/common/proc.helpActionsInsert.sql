print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpActionsInsert') is not null
begin
    drop procedure dbo.helpActionsInsert
end
go
create procedure dbo.helpActionsInsert
    @SystemID varchar(12)
   ,@ObjectName varchar(32)
   ,@ActionName varchar(32)
as
begin
    set nocount on
    declare @e integer
           ,@AuditID integer
           ,@State char(2)

    set @e = 0
    while @e = 0
    begin
        begin transaction

        select  @AuditID = a.AuditID
               ,@State = a.State
        from    dbo.helpActions a
        where   a.SystemID = @SystemID
        and     a.ObjectName = @ObjectName
        and     a.ActionName = @ActionName

        if @@rowcount > 0
        begin
            set @AuditID = @AuditID + 1

            if @State = 'ac'
            begin
                break
            end

            update  dbo.helpActions
            set     State = 'ac'
                   ,AuditID = @AuditID
            where   SystemID = @SystemID
            and     ObjectName = @ObjectName
            and     ActionName = @ActionName

            set @e = @@error
        end
        else
        begin
            select  @AuditID = max(a.AuditID)
            from    dbo.helpActions a
            where   a.SystemID = @SystemID
            and     a.ObjectName = @ObjectName
            and     a.ActionName = @ActionName

            set @AuditID = coalesce(@AuditID, 0) + 1

            insert into dbo.helpActions
            (
                SystemID, ObjectName, ActionName,
                State, AuditID
            )
            values
            (
                @SystemID, @ObjectName, @ActionName,
                'ac', @AuditID
            )
            set @e = @@error
        end
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.helpActionsAuditInsert
            @SystemID = @SystemID
           ,@ObjectName = @ObjectName
           ,@ActionName = @ActionName
           ,@AuditID = @AuditID
           ,@ActionType = 'I'
        if @e <> 0
        begin
            break
        end
        break
    end
    if @@trancount > 0
    begin
        if @e <> 0
        begin
            rollback transaction
        end
        else
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

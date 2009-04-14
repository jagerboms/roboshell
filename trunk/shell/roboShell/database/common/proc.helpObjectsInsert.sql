print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpObjectsInsert') is not null
begin
    drop procedure dbo.helpObjectsInsert
end
go
create procedure dbo.helpObjectsInsert
    @SystemID varchar(12)
   ,@ObjectName varchar(32)
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
        from    dbo.helpObjects a
        where   a.SystemID = @SystemID
        and     a.ObjectName = @ObjectName

        if @@rowcount > 0
        begin
            set @AuditID = @AuditID + 1

            if @State = 'ac'
            begin
                break
            end

            update  dbo.helpObjects
            set     State = 'ac'
                   ,AuditID = @AuditID
            where   SystemID = @SystemID
            and     ObjectName = @ObjectName

            set @e = @@error
        end
        else
        begin
            select  @AuditID = max(a.AuditID)
            from    dbo.helpObjectsAudit a
            where   a.SystemID = @SystemID
            and     a.ObjectName = @ObjectName

            set @AuditID = coalesce(@AuditID, 0) + 1

            insert into dbo.helpObjects
            (
                SystemID, ObjectName, State, AuditID
            )
            values
            (
                @SystemID, @ObjectName, 'ac', @AuditID
            )
            set @e = @@error
        end
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.helpObjectsAuditInsert
            @SystemID = @SystemID
           ,@ObjectName = @ObjectName
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
        if @e = 0
        begin
            commit transaction
        end
        else
        begin
            rollback transaction
        end
    end
    return @e
end
go
go

print '.oOo.'
go

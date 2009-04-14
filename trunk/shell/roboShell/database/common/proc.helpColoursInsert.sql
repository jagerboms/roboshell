print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpColoursInsert') is not null
begin
    drop procedure dbo.helpColoursInsert
end
go
create procedure dbo.helpColoursInsert
    @SystemID varchar(12)
   ,@ObjectName varchar(32)
   ,@ColourValue varchar(200)
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
        from    dbo.helpColours a
        where   a.SystemID = @SystemID
        and     a.ObjectName = @ObjectName

        if @@rowcount > 0
        begin
            set @AuditID = @AuditID + 1

            if @State = 'ac'
            begin
                break
            end

            update  dbo.helpColours
            set     State = 'ac'
                   ,AuditID = @AuditID
            where   SystemID = @SystemID
            and     ObjectName = @ObjectName
            and     ColourValue = @ColourValue

            set @e = @@error
        end
        else
        begin
            select  @AuditID = max(a.AuditID)
            from    dbo.helpColours a
            where   a.SystemID = @SystemID
            and     a.ObjectName = @ObjectName

            set @AuditID = coalesce(@AuditID, 0) + 1

            insert into dbo.helpColours
            (
                SystemID, ObjectName, ColourValue,
                State, AuditID
            )
            values
            (
                @SystemID, @ObjectName, @ColourValue,
                'ac', @AuditID
            )
            set @e = @@error
        end
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.helpColoursAuditInsert
            @SystemID = @SystemID
           ,@ObjectName = @ObjectName
           ,@ColourValue = @ColourValue
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

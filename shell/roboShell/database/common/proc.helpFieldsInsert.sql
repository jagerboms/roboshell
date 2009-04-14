print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpFieldsInsert') is not null
begin
    drop procedure dbo.helpFieldsInsert
end
go
create procedure dbo.helpFieldsInsert
    @SystemID varchar(12)
   ,@ObjectName varchar(32)
   ,@FieldName varchar(32)
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
        from    dbo.helpFields a
        where   a.SystemID = @SystemID
        and     a.ObjectName = @ObjectName
        and     a.FieldName = @FieldName

        if @@rowcount > 0
        begin
            set @AuditID = @AuditID + 1

            if @State = 'ac'
            begin
                break
            end

            update  dbo.helpFields
            set     State = 'ac'
                   ,AuditID = @AuditID
            where   SystemID = @SystemID
            and     ObjectName = @ObjectName
            and     FieldName = @FieldName

            set @e = @@error
        end
        else
        begin
            select  @AuditID = max(a.AuditID)
            from    dbo.helpFields a
            where   a.SystemID = @SystemID
            and     a.ObjectName = @ObjectName
            and     a.FieldName = @FieldName

            set @AuditID = coalesce(@AuditID, 0) + 1

            insert into dbo.helpFields
            (
                SystemID, ObjectName, FieldName,
                State, AuditID
            )
            values
            (
                @SystemID, @ObjectName, @FieldName,
                'ac', @AuditID
            )
            set @e = @@error
        end
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.helpFieldsAuditInsert
            @SystemID = @SystemID
           ,@ObjectName = @ObjectName
           ,@FieldName = @FieldName
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

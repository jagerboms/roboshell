print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpFieldsAuditInsert') is not null
begin
    drop procedure dbo.helpFieldsAuditInsert
end
go
create procedure dbo.helpFieldsAuditInsert
    @SystemID varchar(12)
   ,@ObjectName varchar(32)
   ,@FieldName varchar(32)
   ,@AuditID integer
   ,@ActionType char(1)
as
begin
    set nocount on

    insert into dbo.helpFieldsAudit
    (
        SystemID, ObjectName, FieldName,
        AuditID,
        HelpText, State, ActionType
    )
    select  @SystemID
           ,@ObjectName
           ,@FieldName
           ,@AuditID
           ,a.HelpText
           ,a.State
           ,@ActionType
    from    dbo.helpFields a
    where   a.SystemID = @SystemID
    and     a.ObjectName = @ObjectName
    and     a.FieldName = @FieldName

    return @@error
end
go

print '.oOo.'
go

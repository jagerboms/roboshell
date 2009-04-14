print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpColoursAuditInsert') is not null
begin
    drop procedure dbo.helpColoursAuditInsert
end
go
create procedure dbo.helpColoursAuditInsert
    @SystemID varchar(12)
   ,@ObjectName varchar(32)
   ,@ColourValue varchar(200)
   ,@AuditID integer
   ,@ActionType char(1)
as
begin
    set nocount on

    insert into dbo.helpColoursAudit
    (
        SystemID, ObjectName, ColourValue,
        AuditID,
        ValueDescription, State, ActionType
    )
    select  @SystemID
           ,@ObjectName
           ,@ColourValue
           ,@AuditID
           ,a.ValueDescription
           ,a.State
           ,@ActionType
    from    dbo.helpColours a
    where   a.SystemID = @SystemID
    and     a.ObjectName = @ObjectName
    and     a.ColourValue = @ColourValue

    return @@error
end
go

print '.oOo.'
go

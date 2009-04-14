print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpObjectsAuditInsert') is not null
begin
    drop procedure dbo.helpObjectsAuditInsert
end
go
create procedure dbo.helpObjectsAuditInsert
    @SystemID varchar(12)
   ,@ObjectName varchar(32)
   ,@AuditID integer
   ,@ActionType char(1)
as
begin
    set nocount on

    insert into dbo.helpObjectsAudit
    (
        SystemID, ObjectName, AuditID,
        HelpText, ColourText, State,
        ActionType
    )
    select  @SystemID
           ,@ObjectName
           ,@AuditID
           ,a.HelpText
           ,a.ColourText
           ,a.State
           ,@ActionType
    from    dbo.helpObjects a
    where   a.SystemID = @SystemID
    and     a.ObjectName = @ObjectName

    return @@error
end
go

print '.oOo.'
go

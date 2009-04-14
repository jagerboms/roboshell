print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpActionsAuditInsert') is not null
begin
    drop procedure dbo.helpActionsAuditInsert
end
go
create procedure dbo.helpActionsAuditInsert
    @SystemID varchar(12)
   ,@ObjectName varchar(32)
   ,@ActionName varchar(32)
   ,@AuditID integer
   ,@ActionType char(1)
as
begin
    set nocount on

    insert into dbo.helpActionsAudit
    (
        SystemID, ObjectName, ActionName,
        AuditID,
        HelpText, State, ActionType
    )
    select  @SystemID
           ,@ObjectName
           ,@ActionName
           ,@AuditID
           ,a.HelpText
           ,a.State
           ,@ActionType
    from    dbo.helpActions a
    where   a.SystemID = @SystemID
    and     a.ObjectName = @ObjectName
    and     a.ActionName = @ActionName

    return @@error
end
go

print '.oOo.'
go

print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpSystemsAuditInsert') is not null
begin
    drop procedure dbo.helpSystemsAuditInsert
end
go
create procedure dbo.helpSystemsAuditInsert
    @SystemID varchar(12)
   ,@AuditID integer
   ,@ActionType char(1)
as
begin
    set nocount on

    insert into dbo.helpSystemsAudit
    (
        SystemID, AuditID,
        SystemName, Copyright, State,
        ActionType
    )
    select  @SystemID
           ,@AuditID
           ,a.SystemName
           ,a.Copyright
           ,a.State
           ,@ActionType
    from    dbo.helpSystems a
    where   a.SystemID = @SystemID

    return @@error
end
go

print '.oOo.'
go

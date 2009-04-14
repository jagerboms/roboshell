print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpActionsAuditGet') is not null
begin
    drop procedure dbo.helpActionsAuditGet
end
go
create procedure dbo.helpActionsAuditGet
    @SystemID varchar(12)
   ,@ObjectName varchar(32)
   ,@ActionName varchar(32)
as
begin
    set nocount on

    select  a.AuditID
           ,v.ValueDescription Action
           ,a.ActionType
           ,a.UserID
           ,a.AuditTime
           ,dbo.gnFirstLineGet(a.HelpText) Description
           ,t.ValueDescription State
    from    dbo.helpActionsAudit a
    join    dbo.shlTableValues v
    on      v.TableName = 'default'
    and     v.ColumnName = 'ActionType'
    and     v.ColumnValue = a.ActionType
    join    dbo.shlTableValues t
    on      t.TableName = 'default'
    and     t.ColumnName = 'State'
    and     t.ColumnValue = a.State
    where   a.SystemID = @SystemID
    and     a.ObjectName = @ObjectName
    and     a.ActionName = @ActionName
    order by a.AuditID
end
go

print '.oOo.'
go

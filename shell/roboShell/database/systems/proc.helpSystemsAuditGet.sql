print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpSystemsAuditGet') is not null
begin
    drop procedure dbo.helpSystemsAuditGet
end
go
create procedure dbo.helpSystemsAuditGet
    @SystemID varchar(12)
as
begin
    set nocount on

    select  a.AuditID
           ,v.ValueDescription Action
           ,a.ActionType
           ,a.UserID
           ,a.AuditTime
           ,a.SystemName
           ,a.Copyright
           ,t.ValueDescription State
    from    dbo.helpSystemsAudit a
    join    dbo.shlTableValues v
    on      v.TableName = 'default'
    and     v.ColumnName = 'ActionType'
    and     v.ColumnValue = a.ActionType
    join    dbo.shlTableValues t
    on      t.TableName = 'default'
    and     t.ColumnName = 'State'
    and     t.ColumnValue = a.State
    where   a.SystemID = @SystemID
    order by a.AuditID
end
go

print '.oOo.'
go

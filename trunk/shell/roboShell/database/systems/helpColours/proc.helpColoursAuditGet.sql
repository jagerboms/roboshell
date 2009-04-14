print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpColoursAuditGet') is not null
begin
    drop procedure dbo.helpColoursAuditGet
end
go
create procedure dbo.helpColoursAuditGet
    @SystemID varchar(12)
   ,@ObjectName varchar(32)
   ,@ColourValue varchar(200)
as
begin
    set nocount on

    select  a.AuditID
           ,v.ValueDescription Action
           ,a.ActionType
           ,a.UserID
           ,a.AuditTime
           ,a.ValueDescription
           ,t.ValueDescription State
    from    dbo.helpColoursAudit a
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
    and     a.ColourValue = @ColourValue
    order by a.AuditID
end
go

print '.oOo.'
go

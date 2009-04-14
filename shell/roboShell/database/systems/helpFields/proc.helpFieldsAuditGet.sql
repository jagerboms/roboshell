print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpFieldsAuditGet') is not null
begin
    drop procedure dbo.helpFieldsAuditGet
end
go
create procedure dbo.helpFieldsAuditGet
    @SystemID varchar(12)
   ,@ObjectName varchar(32)
   ,@FieldName varchar(32)
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
    from    dbo.helpFieldsAudit a
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
    and     a.FieldName = @FieldName
    order by a.AuditID
end
go

print '.oOo.'
go

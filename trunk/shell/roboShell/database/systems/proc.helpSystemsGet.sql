print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpSystemsGet') is not null
begin
    drop procedure dbo.helpSystemsGet
end
go
create procedure dbo.helpSystemsGet
    @pSystemID varchar(12) = null
as
begin
    set nocount on

    select  a.SystemID
           ,a.SystemName
           ,a.Copyright
           ,a.State
           ,v.ValueDescription StateName
           ,a.AuditID
    from    dbo.helpSystems a
    join    dbo.shlTableValues v
    on      v.TableName = 'default'
    and     v.ColumnName = 'State'
    and     v.ColumnValue = a.State
    where   a.SystemID = coalesce(@pSystemID, a.SystemID)
    order by a.SystemID
end
go

print '.oOo.'
go

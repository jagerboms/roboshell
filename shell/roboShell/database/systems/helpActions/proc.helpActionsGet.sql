print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpActionsGet') is not null
begin
    drop procedure dbo.helpActionsGet
end
go
create procedure dbo.helpActionsGet
    @pSystemID varchar(12) = null
   ,@pObjectName varchar(32) = null
   ,@pActionName varchar(32) = null
as
begin
    set nocount on

    select  a.SystemID
           ,a.ObjectName
           ,a.ActionName
           ,dbo.gnFirstLineGet(coalesce(a.HelpText, a1.HelpText)) Description
           ,coalesce(a.HelpText, a1.HelpText) HelpText
           ,case @pSystemID when 'default' then 'ft' else 
                case a.State when 'dl' then 'dl' else 
                    case when a1.ObjectName is null then 'ac' else 'sh' end end end State
           ,v.ValueDescription StateName
           ,a.AuditID
    from    dbo.helpActions a
    left join dbo.helpActions a1
    on      a1.SystemID = 'default'
    and     a1.ObjectName = a.ObjectName
    and     a1.ActionName = a.ActionName
    and     a1.State = 'ac'
    join    dbo.shlTableValues v
    on      v.TableName = 'default'
    and     v.ColumnName = 'State'
    and     v.ColumnValue = a.State
    where   a.SystemID = @pSystemID
    and     a.ObjectName = coalesce(@pObjectName, a.ObjectName)
    and     a.ActionName = coalesce(@pActionName, a.ActionName)
    order by a.SystemID
           ,a.ObjectName
           ,a.ActionName
end
go

print '.oOo.'
go

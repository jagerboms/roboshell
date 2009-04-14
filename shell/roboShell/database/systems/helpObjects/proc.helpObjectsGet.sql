print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpObjectsGet') is not null
begin
    drop procedure dbo.helpObjectsGet
end
go
create procedure dbo.helpObjectsGet
    @pSystemID varchar(12)
   ,@pObjectName varchar(32) = null
as
begin
    set nocount on

    select  a.SystemID
           ,a.ObjectName
           ,dbo.gnFirstLineGet(coalesce(a.HelpText, o1.HelpText)) Description
           ,coalesce(a.HelpText, o1.HelpText) HelpText
           ,dbo.gnFirstLineGet(coalesce(a.ColourText, o1.ColourText)) Colour
           ,coalesce(a.ColourText, o1.ColourText) ColourText
           ,case @pSystemID when 'default' then 'ft' else 
                case a.State when 'dl' then 'dl' else 
                    case when o1.ObjectName is null then 'ac' else 'sh' end end end State
           ,v.ValueDescription StateName
           ,a.AuditID
    from    dbo.helpObjects a
    left join dbo.helpObjects o1
    on      o1.SystemID = 'default'
    and     o1.ObjectName = a.ObjectName
    and     o1.State = 'ac'
    and     a.SystemID <> 'default'
    join    dbo.shlTableValues v
    on      v.TableName = 'default'
    and     v.ColumnName = 'State'
    and     v.ColumnValue = a.State
    where   a.SystemID = @pSystemID
    and     a.ObjectName = coalesce(@pObjectName, a.ObjectName)
    order by a.SystemID
           ,a.ObjectName
end
go

print '.oOo.'
go

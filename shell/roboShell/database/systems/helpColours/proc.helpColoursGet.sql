print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpColoursGet') is not null
begin
    drop procedure dbo.helpColoursGet
end
go
create procedure dbo.helpColoursGet
    @pSystemID varchar(12)
   ,@pObjectName varchar(32) = null
   ,@pColourValue varchar(200) = null
as
begin
    set nocount on

    select  a.SystemID
           ,a.ObjectName
           ,a.ColourValue
           ,coalesce(a.ValueDescription, c1.ValueDescription) ValueDescription
           ,case @pSystemID when 'default' then 'ft' else 
                case a.State when 'dl' then 'dl' else 
                    case when c1.ObjectName is null then 'ac' else 'sh' end end end State
           ,v.ValueDescription StateName
           ,a.AuditID
    from    dbo.helpColours a
    left join dbo.helpColours c1
    on      c1.SystemID = 'default'
    and     c1.ObjectName = a.ObjectName
    and     c1.ColourValue = a.ColourValue
    and     c1.State = 'ac'
    join    dbo.shlTableValues v
    on      v.TableName = 'default'
    and     v.ColumnName = 'State'
    and     v.ColumnValue = a.State
    where   a.SystemID = @pSystemID
    and     a.ObjectName = coalesce(@pObjectName, a.ObjectName)
    and     a.ColourValue = coalesce(@pColourValue, a.ColourValue)
    order by a.SystemID
           ,a.ObjectName
           ,a.ColourValue
end
go

print '.oOo.'
go

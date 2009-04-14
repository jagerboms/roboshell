print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpFieldsGet') is not null
begin
    drop procedure dbo.helpFieldsGet
end
go
create procedure dbo.helpFieldsGet
    @pSystemID varchar(12)
   ,@pObjectName varchar(32) = null
   ,@pFieldName varchar(32) = null
as
begin
    set nocount on

    select  a.SystemID
           ,a.ObjectName
           ,a.FieldName
           ,dbo.gnFirstLineGet(coalesce(a.HelpText, f1.HelpText)) Description
           ,coalesce(a.HelpText, f1.HelpText) HelpText
           ,case @pSystemID when 'default' then 'ft' else 
                case a.State when 'dl' then 'dl' else 
                    case when f1.ObjectName is null then 'ac' else 'sh' end end end State
           ,v.ValueDescription StateName
           ,a.AuditID
    from    dbo.helpFields a
    left join dbo.helpFields f1
    on      f1.SystemID = 'default'
    and     f1.ObjectName = a.ObjectName
    and     f1.FieldName = a.FieldName
    and     f1.State = 'ac'
    join    dbo.shlTableValues v
    on      v.TableName = 'default'
    and     v.ColumnName = 'State'
    and     v.ColumnValue = a.State
    where   a.SystemID = @pSystemID
    and     a.ObjectName = coalesce(@pObjectName, a.ObjectName)
    and     a.FieldName = coalesce(@pFieldName, a.FieldName)
    order by a.SystemID
           ,a.ObjectName
           ,a.FieldName
end
go

print '.oOo.'
go

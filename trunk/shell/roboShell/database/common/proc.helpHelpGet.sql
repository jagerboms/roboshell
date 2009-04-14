print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if object_id('dbo.helpHelpGet') is not null
begin
    drop procedure dbo.helpHelpGet
end
go
create procedure dbo.helpHelpGet
    @SystemID varchar(12)
as
begin
    set nocount on

  -- Objects
    select  o.ObjectName
           ,coalesce(o.HelpText, o1.HelpText) HelpText
           ,coalesce(o.ColourText, o1.ColourText) ColourText
           ,case when s.ObjectName is null then 'N' else 'Y' end Shell
    from    dbo.helpObjects o
    left join dbo.helpObjects o1
    on      o1.SystemID = 'default'
    and     o1.ObjectName = o.ObjectName
    and     o1.State = 'ac'
    left join dbo.helpSystemModules() s
    on      s.ObjectName = o.ObjectName
    where   o.SystemID = @SystemID
    and     o.State = 'ac'

  -- Fields
    select  f.ObjectName
           ,f.FieldName
           ,coalesce(f.HelpText, f1.HelpText) HelpText
    from    dbo.helpFields f
    left join dbo.helpFields f1
    on      f1.SystemID = 'default'
    and     f1.ObjectName = f.ObjectName
    and     f1.FieldName = f.FieldName
    and     f1.State = 'ac'
    where   f.SystemID = @SystemID
    and     f.State = 'ac'

  -- Actions
    select  a.ObjectName
           ,a.ActionName
           ,coalesce(a.HelpText, a1.HelpText) HelpText
    from    dbo.helpActions a
    left join dbo.helpActions a1
    on      a1.SystemID = 'default'
    and     a1.ObjectName = a.ObjectName
    and     a1.ActionName = a.ActionName
    and     a1.State = 'ac'
    where   a.SystemID = @SystemID
    and     a.State = 'ac'

  -- Colours
    select  c.ObjectName
           ,c.ColourValue
           ,c.ValueDescription
    from    dbo.helpColours c
    left join dbo.helpColours c1
    on      c1.SystemID = 'default'
    and     c1.ObjectName = c.ObjectName
    and     c1.ValueDescription = c.ValueDescription
    and     c1.State = 'ac'
    where   c.SystemID = @SystemID
    and     c.State = 'ac'

  -- Properties
    select  s.Copyright
    from    dbo.helpSystems s
    where   s.SystemID = @SystemID
    and     s.State = 'ac'

    return 0
end
go

print '.oOo.'
go

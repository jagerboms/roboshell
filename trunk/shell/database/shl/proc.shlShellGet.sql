print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if object_id('dbo.shlShellGet') is not null
begin
    drop procedure dbo.shlShellGet
end
go

create procedure dbo.shlShellGet
as
begin
    set nocount on
    declare @e integer

    set @e = 0
    while @e = 0
    begin
        declare @temp table
        (
            ProcessName char(32) not null
           ,ObjectName  char(32) not null
        )

        set @e = @@error
        if @e <> 0
        begin
            break
        end

        declare @Actions table
        (
            ObjectName char(32) not null
           ,ActionName char(32) not null
        )

        set @e = @@error
        if @e <> 0
        begin
            break
        end

        insert into @temp
        select  n.ProcessName
               ,n.ObjectName
        from    dbo.shlUserProcessGet() n

        set @e = @@error
        if @e <> 0
        begin
            break
        end

  -- Include Menu objects. 
  --   This is necessary as menus are not referenced by processes
  --   but rather call processes.
        insert into @temp
        select  a.Process
               ,o.ObjectName
        from    dbo.shlObjects o
        join    dbo.shlActions a
        on      o.ObjectName = a.ObjectName
        join    @temp t
        on      t.ProcessName = a.Process
        where   o.ObjectType = 'Menu'

        set @e = @@error
        if @e <> 0
        begin
            break
        end

  -- Include menus that only have sub-items.
        insert into @temp
        select  distinct
                ''
               ,a.ObjectName
        from    dbo.shlActions a
        join    @temp t
        on      t.ObjectName = a.Parent
        left join @temp t2
        on      t2.ObjectName = a.ObjectName
        where   t2.ObjectName is null
        and     a.MenuType = 'S'

        set @e = @@error
        if @e <> 0
        begin
            break
        end

  -- Get user action list
        insert into @Actions
        select  distinct
                a.ObjectName
               ,a.ActionName
        from    dbo.shlActions a
        join    @temp t
        on      t.ObjectName = a.ObjectName
        where  (a.Process is null
            and a.ProcessField is null
            and coalesce(a.Parent, '') = '')
        or      a.Process in
                (
                    select  n.ProcessName
                    from    @temp n
                )
        or      exists
                (
                    select  'a'
                    from    dbo.shlActionProcessRules r
                    join    @temp n
                    on      n.ProcessName = r.Process
                    where   r.ObjectName = a.ObjectName
                    and     r.ActionName = a.ActionName
                    and     r.Process is not null
                )

        set @e = @@error
        if @e <> 0
        begin
            break
        end

  -- Variables
        select  v.VariableID
               ,v.VariableValue
        from    dbo.shlVariables v
        where   v.ShellUse = 'Y'

  -- Processes
        select  lower(n.ProcessName) ProcessName
               ,lower(n.SuccessProcess) SuccessProcess
               ,lower(n.FailProcess) FailProcess
               ,n.ConfirmMsg
               ,n.UpdateParent
               ,lower(n.ObjectName) ObjectName
               ,n.LoadVariables
        from    dbo.shlUserProcessGet() n
        order by n.ProcessName
  
  -- Objects
        select  distinct
                lower(o.ObjectName) ObjectName
               ,o.ObjectType
        from    dbo.shlObjects o
        join    @temp t
        on      t.ObjectName = o.ObjectName
        order by 1

  -- Properties
        select  distinct
                lower(p.ObjectName) ObjectName
               ,p.PropertyType
               ,p.PropertyName
               ,'N' UserSpecific
               ,p.Value
        from    dbo.shlProperties p
        join    @temp t
        on      t.ObjectName = p.ObjectName
        left join dbo.shlUserProperties u
        on      u.ObjectName = p.ObjectName
        and     u.PropertyName = p.PropertyName
        and     u.UserName = suser_sname()
        where   u.ObjectName is null
        union
        select  distinct
                lower(p.ObjectName)
               ,'u'
               ,p.PropertyName
               ,'Y'
               ,p.Value
        from    dbo.shlUserProperties p
        join    @temp t
        on      t.ObjectName = p.ObjectName
        where   p.UserName = suser_sname()
        order by 1
               ,p.PropertyName

  -- Parameters
        select  distinct
                lower(p.ObjectName) ObjectName
               ,p.Sequence
               ,p.ParameterName
               ,p.IsInput Input
               ,p.IsOutput Output
               ,p.ValueType
               ,p.Width
               ,p.Value
               ,p.Type
        from    dbo.shlParameters p
        join    @temp t
        on      t.ObjectName = p.ObjectName
        order by 1
               ,p.Sequence

  -- Actions
        select  lower(a.ObjectName) ObjectName
               ,a.Sequence
               ,a.ActionName
               ,coalesce(lower(a.Process), case
                    when MenuType <> 'N' or CloseObject <> 'N' then 'null'
                    else ''
                end) Process
               ,a.RowBased
               ,a.Validate
               ,a.CloseObject
               ,a.IsDblClick
               ,a.IsButton
               ,a.ImageFile
               ,a.ToolTip
               ,a.MenuType
               ,a.MenuText
               ,lower(a.Parent) Parent
               ,a.IsKey
               ,a.KeyCode
               ,a.Shift
               ,a.FieldName
               ,a.ProcessField
               ,a.LinkedParam
               ,a.ParamValue
        from    dbo.shlActions a
        join    @Actions t
        on      t.ObjectName = a.ObjectName
        and     t.ActionName = a.ActionName
        order by 1
               ,a.Sequence

  -- ActionRules
        select  distinct
                lower(r.ObjectName) ObjectName
               ,r.ActionName
               ,r.RuleID
               ,r.RuleName
               ,r.FieldName
               ,r.ValidationType
               ,r.Value
        from    dbo.shlActionRules r
        join    @Actions t
        on      t.ObjectName = r.ObjectName
        and     t.ActionName = r.ActionName
        order by 1
               ,r.ActionName
               ,r.RuleName
               ,r.RuleID

  -- Fields
        select  distinct
                lower(f.ObjectName) ObjectName
               ,f.Sequence
               ,f.FieldName
               ,f.Label
               ,f.Width
               ,f.DisplayType
               ,lower(f.FillProcess) FillProcess
               ,f.TextField
               ,f.ValueField
               ,f.LinkColumn
               ,f.LinkField
               ,f.DisplayWidth
               ,coalesce(f.DisplayHeight, 1) DisplayHeight
               ,0 Decimals
               ,f.Format
               ,f.IsPrimary "Primary"
               ,f.Justify
               ,f.Required
               ,f.Locate
               ,f.ValueType
               ,f.HelpText
               ,f.LabelWidth
               ,coalesce(f.Decimals, -1) "Decimal"
               ,f.NullText
               ,f.Container
        from    dbo.shlFields f
        join    @temp t
        on      t.ObjectName = f.ObjectName
        order by 1
               ,f.Sequence

  -- Validations
        select  distinct
                lower(v.ObjectName) ObjectName
               ,v.ValidationName
               ,v.FieldName
               ,v.ValidationType
               ,v.ValueType
               ,lower(v.Process) Process
               ,v.Value
               ,v.Message
               ,v.ReturnParameter
        from    dbo.shlValidations v
        join    @temp t
        on      t.ObjectName = v.ObjectName
        where   v.Process in
        (
            select  n.ProcessName
            from    @temp n
        )
        or     v.ValueType <> 'P'
        order by 1
               ,v.ValidationName

  -- ValidationRules
        select  distinct
                lower(r.ObjectName) ObjectName
               ,r.ValidationName
               ,r.FieldName
        from    dbo.shlValidationRules r
        join    dbo.shlValidations v
        on      v.ObjectName = r.ObjectName
        and     v.ValidationName = r.ValidationName
        join    @temp t
        on      t.ObjectName = r.ObjectName
        where   v.Process in
        (
            select  n.ProcessName
            from    @temp n
        )
        or     v.ValueType <> 'P'
        order by 1
               ,r.ValidationName

  -- ActionProcessRules
        select  distinct
                lower(r.ObjectName) ObjectName
               ,r.ActionName
               ,r.Value
               ,lower(r.Process) Process
        from    dbo.shlActionProcessRules r
        join    @Actions t
        on      t.ObjectName = r.ObjectName
        and     t.ActionName = r.ActionName
        order by 1
               ,r.ActionName
               ,r.Value

        select  1.1 ShellVersion
        break
    end
    return @e
end
go

print '.oOo.'
go

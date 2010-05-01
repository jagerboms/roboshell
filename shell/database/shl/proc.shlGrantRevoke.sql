print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlGrantRevoke') is not null
begin
    drop procedure dbo.shlGrantRevoke
end
go

create procedure dbo.shlGrantRevoke
as
begin
    set nocount on
    declare @e integer
           ,@vTemp varchar(255)

    set @e = @@error
    while @e = 0
    begin
        create  table #current
        (
            ProcedureName sysname collate database_default not null
           ,Grantee sysname collate database_default not null
           ,GrantDeny char(1) collate database_default null
        )
        set @e = @@error
        if @e <> 0
        begin
            break
        end

        create table #required
        (
            ModuleID varchar(32) collate database_default not null
           ,Grantee sysname collate database_default not null
           ,GrantDeny char(1) collate database_default not null
           ,Source char(1) collate database_default not null
        )
        set @e = @@error
        if @e <> 0
        begin
            break
        end

        create table #reqsp
        (
            ProcedureName sysname collate database_default not null
           ,Grantee sysname collate database_default not null
           ,GrantDeny char(1) collate database_default not null
        )
        set @e = @@error
        if @e <> 0
        begin
            break
        end

        create  table #outp
        (
            OrderBy   char(1) collate database_default not null
           ,Action    varchar(255) collate database_default not null
        )
        set @e = @@error
        if @e <> 0
        begin
            break
        end

    -- collect current database permission settings

        insert  #current
        select  object_name(p.id)
               ,user_name(p.uid)
               ,substring(v2.name, 1, 1)
        from    dbo.sysprotects p
        join    master.dbo.spt_values v1
        on      v1.number = p.action
        join    master.dbo.spt_values v2
        on      v2.number = p.ProtectType
        join    dbo.sysusers s
        on      p.uid = s.uid
        where   s.uid not in
        (
            select  m.memberuid
            from    dbo.sysmembers m
            join    dbo.sysusers u
            on      m.groupuid = u.uid
            where   u.name = 'db_owner'
            and     m.memberuid = s.uid
        )
        and     v1.type  = 'T'
        and     v1.name = 'Execute'
        and     v2.type = 'T'
        and     p.id <> 0

        set @e = @@error
        if @e <> 0
        begin
            break
        end

  -- determine the required security settings

        insert  #required
        select  distinct
                m.ModuleID
               ,u.UserName
               ,u.GrantDeny
               ,'A'
        from    dbo.shlModuleUsers u
        join    dbo.shlModules m
        on      u.ModuleID = m.ModuleID
        join    dbo.sysusers s
        on      s.name = u.UserName
        where   s.uid not in
        (
            select  m.memberuid
            from    dbo.sysmembers m
            join    dbo.sysusers u
            on      m.groupuid = u.uid
            where   u.name = 'db_owner'
            and     m.memberuid = s.uid
        )
        and     s.name <> 'db_owner'
        set @e = @@error
        if @e <> 0
        begin
            break
        end

        while @e = 0
        begin
            insert  #required
            select  distinct
                    m.ModuleID
                   ,t.Grantee
                   ,t.GrantDeny
                   ,'B'
            from    #required t
            join    dbo.shlModuleOwners o
            on      o.OwnerModule = t.ModuleID
            join    dbo.shlModules m
            on      m.ModuleID = o.ModuleID
            where   o.ModuleID not in
            (
                select  n.ModuleID
                from    #required n
                where   n.GrantDeny = t.GrantDeny
                and     n.Grantee = t.Grantee
                and     n.ModuleID = o.ModuleID
            )
            and not exists
            (
                select  'a'
                from    #required n
                where   n.ModuleID = o.ModuleID
                and     n.Grantee = t.Grantee
                and     n.Source = 'A'
            )
            if @@rowcount = 0
            begin
                break
            end
        end
        if @e <> 0 
        begin
            break
        end

    -- remove the records we are not interested in.

        delete
        from    #required
        where   GrantDeny not in ('G', 'D')

        set @e = @@error
        if @e <> 0 
        begin
            break
        end

        insert into #reqsp
        select  distinct
                o.Value
               ,r.Grantee
               ,r.GrantDeny
        from    #required r
        join    dbo.shlModuleProcesses mp
        on      r.ModuleID = mp.ModuleID
        join    dbo.shlProcesses p
        on      p.ProcessName = mp.ProcessName
        join    dbo.shlProperties o
        on      o.ObjectName = p.ObjectName
        and     o.PropertyType = 'df'
        and     o.PropertyName = 'procname'

        set @e = @@error
        if @e <> 0 
        begin
            break
        end

        insert into #reqsp
        select  distinct
                m.ProcedureName
               ,r.Grantee
               ,r.GrantDeny
        from    #required r
        join    dbo.shlModuleProcedures m
        on      r.ModuleID = m.ModuleID
        where   r.GrantDeny = 'G'
        and     not exists
        (
            select  'a'
            from    #reqsp q
            where   q.ProcedureName = m.ProcedureName
            and     q.Grantee = r.Grantee
        )

        set @e = @@error
        if @e <> 0 
        begin
            break
        end

        insert into #reqsp
        select  distinct
                o.Value
               ,r.Grantee
               ,r.GrantDeny
        from    #required r
        join    dbo.shlModuleProcesses mp
        on      r.ModuleID = mp.ModuleID
        join    dbo.shlProcesses p1
        on      p1.ProcessName = mp.ProcessName
        join    dbo.shlFields f
        on      f.ObjectName = p1.ObjectName
        join    dbo.shlProcesses p2
        on      f.FillProcess = p2.ProcessName
        join    dbo.shlProperties o
        on      o.ObjectName = p2.ObjectName
        and     o.PropertyType = 'df'
        and     o.PropertyName = 'procname'
        where   r.GrantDeny = 'G'
        and     not exists
        (
            select  'a'
            from    #reqsp q
            where   q.ProcedureName = o.Value
            and     q.Grantee = r.Grantee
        )

        set @e = @@error
        if @e <> 0 
        begin
            break
        end

    -- revoke permissions from currently granted procedures that are not required.

        insert  #outp
        select  'b'
               ,Action = 'revoke execute on ' +
                rtrim(p.ProcedureName) + ' to "' + rtrim(p.Grantee) + '"'
        from    #current p
        where   not exists
        (
            select  'a'
            from    #reqsp t
            where   p.ProcedureName = t.ProcedureName
            and     p.Grantee = t.Grantee
        )
        set @e = @@error
        if @e <> 0 
        begin
            break
        end

    -- grant/deny permissions to procedures that are required but not currently granted.

        insert  #outp
        select  'c'
               ,Action = case t.GrantDeny when 'G' then 'grant' else 'deny' end
                + ' execute on ' +
                rtrim(t.ProcedureName) + ' to "' + rtrim(t.Grantee) + '"'
        from    #reqsp t
        where   not exists
        (
            select  'a'
            from    #current p
            where   p.Grantee = t.Grantee
            and     p.ProcedureName = t.ProcedureName
            and     p.GrantDeny = t.GrantDeny
        )
        and     object_id(t.ProcedureName) is not null
        set @e = @@error
        if @e <> 0 
        begin
            break
        end
        break
    end
    if @e = 0
    begin
        while 1 = 1
        begin
            select  top 1
                    @vTemp = o.Action
            from    #outp o
            order by o.OrderBy, o.Action desc
            if @@rowcount = 0
            begin
                break
            end

            execute (@vTemp)

            delete  o
            from    #outp o
            where   @vTemp = o.Action
        end
    end
end
go

print '.oOo.'
go

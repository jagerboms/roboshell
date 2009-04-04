print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if object_id('dbo.shlModulePermGet') is not null
begin
    drop function dbo.shlModulePermGet
end
go
create function dbo.shlModulePermGet()
returns @rs table
(
    ModuleID varchar(32) collate database_default not null
   ,GrantDeny char(1) collate database_default not null
)
as
begin
    declare @e integer
           ,@c integer

    declare @tt table
    (
        ModuleID varchar(32) collate database_default not null
       ,UserName sysname collate database_default not null
       ,GrantDeny char(1) collate database_default not null
       ,Source char(1) collate database_default not null
    )

    set @e = 0
    while @e = 0
    begin

  -- retrieve top level permissions for current user
  -- including permissions for the groups and roles they are members of

        insert into @tt
        select  distinct
                m.ModuleID
               ,mu.UserName
               ,mu.GrantDeny
               ,'A'
        from    dbo.shlModuleUsers mu
        join    dbo.shlModules m
        on      mu.ModuleID = m.ModuleID
        join    dbo.sysusers u
        on      u.name = mu.UserName 
        where  (dbo.shlUserNameGet(u.name) = dbo.shlUserNameGet(suser_sname()))
        or     (is_member(u.name) = 1 and (u.isntgroup = 1 or u.issqlrole = 1))
        or     (u.issqlrole = 1 and exists
        (
            select  'a'
            from    dbo.sysmembers m
            join    dbo.sysusers u2
            on      u2.uid = m.memberuid
            where   u.uid = m.groupuid
            and     is_member(dbo.shlUserNameAddDomain(u2.name)) = 1
        ))

        select  @e = @@error
               ,@c = @@rowcount
        if @e <> 0
        begin
            break
        end

    -- child modules inherit the permissions of their parent
    -- loop until all nested levels are satisfied.

        while @e = 0 and @c <> 0
        begin
            insert into @tt
            select  distinct
                    m.ModuleID
                   ,t.UserName
                   ,t.GrantDeny
                   ,'B'
            from    @tt t
            join    dbo.shlModuleOwners o
            on      o.OwnerModule = t.ModuleID
            join    dbo.shlModules m
            on      m.ModuleID = o.ModuleID
            where   o.ModuleID not in
            (
                select  n.ModuleID
                from    @tt n
                where   n.GrantDeny = t.GrantDeny
                and     n.ModuleID = o.ModuleID
            )
            and not exists
            (
                select  'a'
                from    @tt n
                where   n.ModuleID = o.ModuleID
                and     n.UserName = t.UserName
                and     n.Source = 'A'
            )
            select  @e = @@error
                   ,@c = @@rowcount
        end
        if @e <> 0 
        begin
            break
        end

    -- remove any other user permissions if deny exists

        delete  t
        from    @tt t
        where   t.ModuleID in
        (
            select  n.ModuleID
            from    @tt n
            where   n.GrantDeny = 'D'
            and     t.ModuleID = n.ModuleID
        )
        and     t.GrantDeny <> 'D'

    -- remove any other user permissions if grant exists

        delete  t
        from    @tt t
        where   t.ModuleID in
        (
            select  n.ModuleID
            from    @tt n
            where   n.GrantDeny = 'G'
            and     t.ModuleID = n.ModuleID
        )
        and     t.GrantDeny <> 'G'

    -- results are returned

        insert into @rs
        select  distinct
                t.ModuleID
               ,t.GrantDeny
        from    @tt t

        break
    end
    return
end
go

print '.oOo.'
go

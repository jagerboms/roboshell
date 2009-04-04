print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlUserPermGet') is not null
begin
    drop procedure dbo.shlUserPermGet
end
go

create Procedure dbo.shlUserPermGet
    @UserName sysname
as
begin
    set nocount on
    declare @e integer
           ,@count integer

    set @e = 0
    while @e = 0
    begin
        create table #temp
        (
            Parent char(32) not null
           ,ModuleID char(32) not null
           ,Description varchar(50) null
           ,Type char(1) null
           ,Source char(1) null
        )
        set @e = @@error
        if @e <> 0
        begin
            break
        end

        insert into #temp
        select  lower(coalesce(o.OwnerModule, m.ModuleID))
               ,lower(m.ModuleID)
               ,m.Description
               ,u.GrantDeny
               ,case when u.GrantDeny is null then 'I' else 'A' end
        from    dbo.shlModules m
        left join dbo.shlModuleOwners o
        on      o.ModuleID = m.ModuleID
        left join dbo.shlModuleUsers u
        on      u.ModuleID = m.ModuleID
        and     u.UserName = @UserName

        select  @e = @@error
               ,@count = @@rowcount
        if @e <> 0
        begin
            break
        end

        while @e = 0 and @count <> 0
        begin
            update  t
            set     t.Type = p.Type
                   ,t.Source = 'I'
            from    #temp t
            join    #temp p
            on      t.Parent = p.ModuleID
            and     p.Type is not null
            where   t.Type is null

            select  @e = @@error
                   ,@count = @@rowcount
        end
        if @e <> 0 
        begin
            break
        end

        select  t.Parent
               ,t.ModuleID
               ,Description = case t.ModuleID when 'base' then @UserName else t.Description end
               ,Type = coalesce(t.Type, 'N') + coalesce(t.Source, 'I')
        from    #temp t

        break
    end
    return @e
end
go

print '.oOo.'
go

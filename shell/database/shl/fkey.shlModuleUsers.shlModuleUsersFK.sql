print '----------'
print '-- Pets --'
print '----------'
set nocount on
go
if (
    select  count(*)
    from    dbo.sysforeignkeys k
    join
    (
        select  1 keyno, 'ModuleID' lkey, 'ModuleID' fkey
    ) x
    on      x.keyno = k.keyno
    and     x.lkey = col_name(k.fkeyid, k.fkey)
    and     x.fkey = col_name(k.rkeyid, k.rkey)
    where   k.constid = object_id('shlModuleUsersFK')
    and     k.fkeyid = object_id('shlModuleUsers')
    and     k.rkeyid = object_id('shlModules')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('shlModuleUsersFK')
        and     k.fkeyid = object_id('shlModuleUsers')
        and     k.rkeyid = object_id('shlModules')
    ) =  1
) <>  1
begin
    if object_id('shlModuleUsersFK') is not null
    begin
        print 'changing foreign key ''shlModuleUsersFK'''
        alter table dbo.shlModuleUsers drop constraint shlModuleUsersFK
    end
    else
    begin
        print 'creating foreign key ''shlModuleUsersFK'''
    end
    alter table dbo.shlModuleUsers add constraint shlModuleUsersFK
    foreign key (ModuleID) references dbo.shlModules(ModuleID)
end
go

print '.oOo.'
go

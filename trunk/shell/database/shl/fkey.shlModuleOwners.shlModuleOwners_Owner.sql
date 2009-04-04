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
        select  1 keyno, 'OwnerModule' lkey, 'ModuleID' fkey
    ) x
    on      x.keyno = k.keyno
    and     x.lkey = col_name(k.fkeyid, k.fkey)
    and     x.fkey = col_name(k.rkeyid, k.rkey)
    where   k.constid = object_id('shlModuleOwners_Owner')
    and     k.fkeyid = object_id('shlModuleOwners')
    and     k.rkeyid = object_id('shlModules')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('shlModuleOwners_Owner')
        and     k.fkeyid = object_id('shlModuleOwners')
        and     k.rkeyid = object_id('shlModules')
    ) =  1
) <>  1
begin
    if object_id('shlModuleOwners_Owner') is not null
    begin
        print 'changing foreign key ''shlModuleOwners_Owner'''
        alter table dbo.shlModuleOwners drop constraint shlModuleOwners_Owner
    end
    else
    begin
        print 'creating foreign key ''shlModuleOwners_Owner'''
    end
    alter table dbo.shlModuleOwners add constraint shlModuleOwners_Owner
    foreign key (OwnerModule) references dbo.shlModules(ModuleID)
end
go

print '.oOo.'
go

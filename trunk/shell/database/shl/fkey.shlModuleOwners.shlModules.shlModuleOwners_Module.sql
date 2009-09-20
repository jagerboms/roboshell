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
    where   k.constid = object_id('shlModuleOwners_Module')
    and     k.fkeyid = object_id('shlModuleOwners')
    and     k.rkeyid = object_id('shlModules')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('shlModuleOwners_Module')
        and     k.fkeyid = object_id('shlModuleOwners')
        and     k.rkeyid = object_id('shlModules')
    ) =  1
) <>  1
begin
    if object_id('shlModuleOwners_Module') is not null
    begin
        print 'changing foreign key ''shlModuleOwners_Module'''
        alter table dbo.shlModuleOwners drop constraint shlModuleOwners_Module
    end
    else
    begin
        print 'creating foreign key ''shlModuleOwners_Module'''
    end
    alter table dbo.shlModuleOwners add constraint shlModuleOwners_Module
    foreign key (ModuleID) references dbo.shlModules(ModuleID)
end
go

print '.oOo.'
go

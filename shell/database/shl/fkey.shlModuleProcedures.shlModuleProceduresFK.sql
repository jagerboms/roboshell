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
    where   k.constid = object_id('shlModuleProceduresFK')
    and     k.fkeyid = object_id('shlModuleProcedures')
    and     k.rkeyid = object_id('shlModules')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('shlModuleProceduresFK')
        and     k.fkeyid = object_id('shlModuleProcedures')
        and     k.rkeyid = object_id('shlModules')
    ) =  1
) <>  1
begin
    if object_id('shlModuleProceduresFK') is not null
    begin
        print 'changing foreign key ''shlModuleProceduresFK'''
        alter table dbo.shlModuleProcedures drop constraint shlModuleProceduresFK
    end
    else
    begin
        print 'creating foreign key ''shlModuleProceduresFK'''
    end
    alter table dbo.shlModuleProcedures add constraint shlModuleProceduresFK
    foreign key (ModuleID) references dbo.shlModules(ModuleID)
end
go

print '.oOo.'
go

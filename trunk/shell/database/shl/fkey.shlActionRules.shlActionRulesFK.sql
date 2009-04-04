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
        select  1 keyno, 'ObjectName' lkey, 'ObjectName' fkey
        union select  2, 'ActionName', 'ActionName'
    ) x
    on      x.keyno = k.keyno
    and     x.lkey = col_name(k.fkeyid, k.fkey)
    and     x.fkey = col_name(k.rkeyid, k.rkey)
    where   k.constid = object_id('shlActionRulesFK')
    and     k.fkeyid = object_id('shlActionRules')
    and     k.rkeyid = object_id('shlActions')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('shlActionRulesFK')
        and     k.fkeyid = object_id('shlActionRules')
        and     k.rkeyid = object_id('shlActions')
    ) =  2
) <>  2
begin
    if object_id('shlActionRulesFK') is not null
    begin
        print 'changing foreign key ''shlActionRulesFK'''
        alter table dbo.shlActionRules drop constraint shlActionRulesFK
    end
    else
    begin
        print 'creating foreign key ''shlActionRulesFK'''
    end
    alter table dbo.shlActionRules add constraint shlActionRulesFK
    foreign key (ObjectName,ActionName) references dbo.shlActions(ObjectName,ActionName)
end
go

print '.oOo.'
go

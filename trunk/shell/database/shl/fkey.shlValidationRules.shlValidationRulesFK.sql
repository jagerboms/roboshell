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
        union select  2, 'ValidationName', 'ValidationName'
    ) x
    on      x.keyno = k.keyno
    and     x.lkey = col_name(k.fkeyid, k.fkey)
    and     x.fkey = col_name(k.rkeyid, k.rkey)
    where   k.constid = object_id('shlValidationRulesFK')
    and     k.fkeyid = object_id('shlValidationRules')
    and     k.rkeyid = object_id('shlValidations')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('shlValidationRulesFK')
        and     k.fkeyid = object_id('shlValidationRules')
        and     k.rkeyid = object_id('shlValidations')
    ) =  2
) <>  2
begin
    if object_id('shlValidationRulesFK') is not null
    begin
        print 'changing foreign key ''shlValidationRulesFK'''
        alter table dbo.shlValidationRules drop constraint shlValidationRulesFK
    end
    else
    begin
        print 'creating foreign key ''shlValidationRulesFK'''
    end
    alter table dbo.shlValidationRules add constraint shlValidationRulesFK
    foreign key (ObjectName,ValidationName) references dbo.shlValidations(ObjectName,ValidationName)
end
go

print '.oOo.'
go

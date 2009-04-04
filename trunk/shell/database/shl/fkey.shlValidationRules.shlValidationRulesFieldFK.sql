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
        union select  2, 'FieldName', 'FieldName'
    ) x
    on      x.keyno = k.keyno
    and     x.lkey = col_name(k.fkeyid, k.fkey)
    and     x.fkey = col_name(k.rkeyid, k.rkey)
    where   k.constid = object_id('shlValidationRulesFieldFK')
    and     k.fkeyid = object_id('shlValidationRules')
    and     k.rkeyid = object_id('shlFields')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('shlValidationRulesFieldFK')
        and     k.fkeyid = object_id('shlValidationRules')
        and     k.rkeyid = object_id('shlFields')
    ) =  2
) <>  2
begin
    if object_id('shlValidationRulesFieldFK') is not null
    begin
        print 'changing foreign key ''shlValidationRulesFieldFK'''
        alter table dbo.shlValidationRules drop constraint shlValidationRulesFieldFK
    end
    else
    begin
        print 'creating foreign key ''shlValidationRulesFieldFK'''
    end
    alter table dbo.shlValidationRules add constraint shlValidationRulesFieldFK
    foreign key (ObjectName,FieldName) references dbo.shlFields(ObjectName,FieldName)
end
go

print '.oOo.'
go

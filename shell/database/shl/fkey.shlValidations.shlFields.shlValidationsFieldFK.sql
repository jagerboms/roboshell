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
    where   k.constid = object_id('shlValidationsFieldFK')
    and     k.fkeyid = object_id('shlValidations')
    and     k.rkeyid = object_id('shlFields')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('shlValidationsFieldFK')
        and     k.fkeyid = object_id('shlValidations')
        and     k.rkeyid = object_id('shlFields')
    ) =  2
) <>  2
begin
    if object_id('shlValidationsFieldFK') is not null
    begin
        print 'changing foreign key ''shlValidationsFieldFK'''
        alter table dbo.shlValidations drop constraint shlValidationsFieldFK
    end
    else
    begin
        print 'creating foreign key ''shlValidationsFieldFK'''
    end
    alter table dbo.shlValidations add constraint shlValidationsFieldFK
    foreign key (ObjectName,FieldName) references dbo.shlFields(ObjectName,FieldName)
end
go

print '.oOo.'
go

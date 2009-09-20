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
    ) x
    on      x.keyno = k.keyno
    and     x.lkey = col_name(k.fkeyid, k.fkey)
    and     x.fkey = col_name(k.rkeyid, k.rkey)
    where   k.constid = object_id('shlPropertiesFK')
    and     k.fkeyid = object_id('shlProperties')
    and     k.rkeyid = object_id('shlObjects')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('shlPropertiesFK')
        and     k.fkeyid = object_id('shlProperties')
        and     k.rkeyid = object_id('shlObjects')
    ) =  1
) <>  1
begin
    if object_id('shlPropertiesFK') is not null
    begin
        print 'changing foreign key ''shlPropertiesFK'''
        alter table dbo.shlProperties drop constraint shlPropertiesFK
    end
    else
    begin
        print 'creating foreign key ''shlPropertiesFK'''
    end
    alter table dbo.shlProperties add constraint shlPropertiesFK
    foreign key (ObjectName) references dbo.shlObjects(ObjectName)
end
go

print '.oOo.'
go

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
    where   k.constid = object_id('shlParametersFK')
    and     k.fkeyid = object_id('shlParameters')
    and     k.rkeyid = object_id('shlObjects')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('shlParametersFK')
        and     k.fkeyid = object_id('shlParameters')
        and     k.rkeyid = object_id('shlObjects')
    ) =  1
) <>  1
begin
    if object_id('shlParametersFK') is not null
    begin
        print 'changing foreign key ''shlParametersFK'''
        alter table dbo.shlParameters drop constraint shlParametersFK
    end
    else
    begin
        print 'creating foreign key ''shlParametersFK'''
    end
    alter table dbo.shlParameters add constraint shlParametersFK
    foreign key (ObjectName) references dbo.shlObjects(ObjectName)
end
go

print '.oOo.'
go

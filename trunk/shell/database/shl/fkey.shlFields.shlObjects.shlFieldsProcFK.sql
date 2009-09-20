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
    where   k.constid = object_id('shlFieldsProcFK')
    and     k.fkeyid = object_id('shlFields')
    and     k.rkeyid = object_id('shlObjects')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('shlFieldsProcFK')
        and     k.fkeyid = object_id('shlFields')
        and     k.rkeyid = object_id('shlObjects')
    ) =  1
) <>  1
begin
    if object_id('shlFieldsProcFK') is not null
    begin
        print 'changing foreign key ''shlFieldsProcFK'''
        alter table dbo.shlFields drop constraint shlFieldsProcFK
    end
    else
    begin
        print 'creating foreign key ''shlFieldsProcFK'''
    end
    alter table dbo.shlFields add constraint shlFieldsProcFK
    foreign key (ObjectName) references dbo.shlObjects(ObjectName)
end
go

print '.oOo.'
go

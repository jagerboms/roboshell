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
    where   k.constid = object_id('shlProcessesFK')
    and     k.fkeyid = object_id('shlProcesses')
    and     k.rkeyid = object_id('shlObjects')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('shlProcessesFK')
        and     k.fkeyid = object_id('shlProcesses')
        and     k.rkeyid = object_id('shlObjects')
    ) =  1
) <>  1
begin
    if object_id('shlProcessesFK') is not null
    begin
        print 'changing foreign key ''shlProcessesFK'''
        alter table dbo.shlProcesses drop constraint shlProcessesFK
    end
    else
    begin
        print 'creating foreign key ''shlProcessesFK'''
    end
    alter table dbo.shlProcesses add constraint shlProcessesFK
    foreign key (ObjectName) references dbo.shlObjects(ObjectName)
end
go

print '.oOo.'
go

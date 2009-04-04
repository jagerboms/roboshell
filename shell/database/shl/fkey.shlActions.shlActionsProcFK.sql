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
        select  1 keyno, 'Process' lkey, 'ProcessName' fkey
    ) x
    on      x.keyno = k.keyno
    and     x.lkey = col_name(k.fkeyid, k.fkey)
    and     x.fkey = col_name(k.rkeyid, k.rkey)
    where   k.constid = object_id('shlActionsProcFK')
    and     k.fkeyid = object_id('shlActions')
    and     k.rkeyid = object_id('shlProcesses')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('shlActionsProcFK')
        and     k.fkeyid = object_id('shlActions')
        and     k.rkeyid = object_id('shlProcesses')
    ) =  1
) <>  1
begin
    if object_id('shlActionsProcFK') is not null
    begin
        print 'changing foreign key ''shlActionsProcFK'''
        alter table dbo.shlActions drop constraint shlActionsProcFK
    end
    else
    begin
        print 'creating foreign key ''shlActionsProcFK'''
    end
    alter table dbo.shlActions add constraint shlActionsProcFK
    foreign key (Process) references dbo.shlProcesses(ProcessName)
end
go

print '.oOo.'
go

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
        select  1 keyno, 'ProcessName' lkey, 'ProcessName' fkey
    ) x
    on      x.keyno = k.keyno
    and     x.lkey = col_name(k.fkeyid, k.fkey)
    and     x.fkey = col_name(k.rkeyid, k.rkey)
    where   k.constid = object_id('shlModuleProcessesProcFK')
    and     k.fkeyid = object_id('shlModuleProcesses')
    and     k.rkeyid = object_id('shlProcesses')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('shlModuleProcessesProcFK')
        and     k.fkeyid = object_id('shlModuleProcesses')
        and     k.rkeyid = object_id('shlProcesses')
    ) =  1
) <>  1
begin
    if object_id('shlModuleProcessesProcFK') is not null
    begin
        print 'changing foreign key ''shlModuleProcessesProcFK'''
        alter table dbo.shlModuleProcesses drop constraint shlModuleProcessesProcFK
    end
    else
    begin
        print 'creating foreign key ''shlModuleProcessesProcFK'''
    end
    alter table dbo.shlModuleProcesses add constraint shlModuleProcessesProcFK
    foreign key (ProcessName) references dbo.shlProcesses(ProcessName)
end
go

print '.oOo.'
go

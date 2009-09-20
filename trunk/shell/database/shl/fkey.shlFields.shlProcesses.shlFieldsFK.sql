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
        select  1 keyno, 'FillProcess' lkey, 'ProcessName' fkey
    ) x
    on      x.keyno = k.keyno
    and     x.lkey = col_name(k.fkeyid, k.fkey)
    and     x.fkey = col_name(k.rkeyid, k.rkey)
    where   k.constid = object_id('shlFieldsFK')
    and     k.fkeyid = object_id('shlFields')
    and     k.rkeyid = object_id('shlProcesses')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('shlFieldsFK')
        and     k.fkeyid = object_id('shlFields')
        and     k.rkeyid = object_id('shlProcesses')
    ) =  1
) <>  1
begin
    if object_id('shlFieldsFK') is not null
    begin
        print 'changing foreign key ''shlFieldsFK'''
        alter table dbo.shlFields drop constraint shlFieldsFK
    end
    else
    begin
        print 'creating foreign key ''shlFieldsFK'''
    end
    alter table dbo.shlFields add constraint shlFieldsFK
    foreign key (FillProcess) references dbo.shlProcesses(ProcessName)
end
go

print '.oOo.'
go

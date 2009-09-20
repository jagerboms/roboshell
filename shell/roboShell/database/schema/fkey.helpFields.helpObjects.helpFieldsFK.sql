print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if (
    select  count(*)
    from    dbo.sysforeignkeys k
    join
    (
        select  1 keyno, 'SystemID' lkey, 'SystemID' fkey
        union select  2, 'ObjectName', 'ObjectName'
    ) x
    on      x.keyno = k.keyno
    and     x.lkey = col_name(k.fkeyid, k.fkey)
    and     x.fkey = col_name(k.rkeyid, k.rkey)
    where   k.constid = object_id('helpFieldsFK')
    and     k.fkeyid = object_id('helpFields')
    and     k.rkeyid = object_id('helpObjects')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('helpFieldsFK')
        and     k.fkeyid = object_id('helpFields')
        and     k.rkeyid = object_id('helpObjects')
    ) =  2
) <>  2
begin
    if object_id('helpFieldsFK') is not null
    begin
        print 'changing foreign key ''helpFieldsFK'''
        alter table dbo.helpFields drop constraint helpFieldsFK
    end
    else
    begin
        print 'creating foreign key ''helpFieldsFK'''
    end
    alter table dbo.helpFields add constraint helpFieldsFK
    foreign key (SystemID,ObjectName) references dbo.helpObjects(SystemID,ObjectName)
end
go

print '.oOo.'
go

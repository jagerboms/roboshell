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
    ) x
    on      x.keyno = k.keyno
    and     x.lkey = col_name(k.fkeyid, k.fkey)
    and     x.fkey = col_name(k.rkeyid, k.rkey)
    where   k.constid = object_id('helpObjectsFK')
    and     k.fkeyid = object_id('helpObjects')
    and     k.rkeyid = object_id('helpSystems')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('helpObjectsFK')
        and     k.fkeyid = object_id('helpObjects')
        and     k.rkeyid = object_id('helpSystems')
    ) =  1
) <>  1
begin
    if object_id('helpObjectsFK') is not null
    begin
        print 'changing foreign key ''helpObjectsFK'''
        alter table dbo.helpObjects drop constraint helpObjectsFK
    end
    else
    begin
        print 'creating foreign key ''helpObjectsFK'''
    end
    alter table dbo.helpObjects add constraint helpObjectsFK
    foreign key (SystemID) references dbo.helpSystems(SystemID)
end
go

print '.oOo.'
go

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
    where   k.constid = object_id('helpColoursFK')
    and     k.fkeyid = object_id('helpColours')
    and     k.rkeyid = object_id('helpObjects')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('helpColoursFK')
        and     k.fkeyid = object_id('helpColours')
        and     k.rkeyid = object_id('helpObjects')
    ) =  2
) <>  2
begin
    if object_id('helpColoursFK') is not null
    begin
        print 'changing foreign key ''helpColoursFK'''
        alter table dbo.helpColours drop constraint helpColoursFK
    end
    else
    begin
        print 'creating foreign key ''helpColoursFK'''
    end
    alter table dbo.helpColours add constraint helpColoursFK
    foreign key (SystemID,ObjectName) references dbo.helpObjects(SystemID,ObjectName)
end
go

print '.oOo.'
go

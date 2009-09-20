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
    where   k.constid = object_id('helpActionsFK')
    and     k.fkeyid = object_id('helpActions')
    and     k.rkeyid = object_id('helpObjects')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('helpActionsFK')
        and     k.fkeyid = object_id('helpActions')
        and     k.rkeyid = object_id('helpObjects')
    ) =  2
) <>  2
begin
    if object_id('helpActionsFK') is not null
    begin
        print 'changing foreign key ''helpActionsFK'''
        alter table dbo.helpActions drop constraint helpActionsFK
    end
    else
    begin
        print 'creating foreign key ''helpActionsFK'''
    end
    alter table dbo.helpActions add constraint helpActionsFK
    foreign key (SystemID,ObjectName) references dbo.helpObjects(SystemID,ObjectName)
end
go

print '.oOo.'
go

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
        union select  2, 'FieldName', 'FieldName'
    ) x
    on      x.keyno = k.keyno
    and     x.lkey = col_name(k.fkeyid, k.fkey)
    and     x.fkey = col_name(k.rkeyid, k.rkey)
    where   k.constid = object_id('shlActionsFieldFK')
    and     k.fkeyid = object_id('shlActions')
    and     k.rkeyid = object_id('shlFields')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('shlActionsFieldFK')
        and     k.fkeyid = object_id('shlActions')
        and     k.rkeyid = object_id('shlFields')
    ) =  2
) <>  2
begin
    if object_id('shlActionsFieldFK') is not null
    begin
        print 'changing foreign key ''shlActionsFieldFK'''
        alter table dbo.shlActions drop constraint shlActionsFieldFK
    end
    else
    begin
        print 'creating foreign key ''shlActionsFieldFK'''
    end
    alter table dbo.shlActions add constraint shlActionsFieldFK
    foreign key (ObjectName,FieldName) references dbo.shlFields(ObjectName,FieldName)
end
go

print '.oOo.'
go

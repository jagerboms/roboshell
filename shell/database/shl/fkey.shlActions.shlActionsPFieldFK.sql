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
        union select  2, 'ProcessField', 'FieldName'
    ) x
    on      x.keyno = k.keyno
    and     x.lkey = col_name(k.fkeyid, k.fkey)
    and     x.fkey = col_name(k.rkeyid, k.rkey)
    where   k.constid = object_id('shlActionsPFieldFK')
    and     k.fkeyid = object_id('shlActions')
    and     k.rkeyid = object_id('shlFields')
    and
    (
        select  count(*)
        from    dbo.sysforeignkeys k
        where   k.constid = object_id('shlActionsPFieldFK')
        and     k.fkeyid = object_id('shlActions')
        and     k.rkeyid = object_id('shlFields')
    ) =  2
) <>  2
begin
    if object_id('shlActionsPFieldFK') is not null
    begin
        print 'changing foreign key ''shlActionsPFieldFK'''
        alter table dbo.shlActions drop constraint shlActionsPFieldFK
    end
    else
    begin
        print 'creating foreign key ''shlActionsPFieldFK'''
    end
    alter table dbo.shlActions add constraint shlActionsPFieldFK
    foreign key (ObjectName,ProcessField) references dbo.shlFields(ObjectName,FieldName)
end
go

print '.oOo.'
go

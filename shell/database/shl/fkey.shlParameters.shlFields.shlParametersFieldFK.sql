print '-----------------------'
print '-- Akuna Care - Pets --'
print '-----------------------'
set nocount on
go
declare @c1 integer, @c2 integer

if object_id('shlParametersFieldFK') is not null
begin
    select  @c1 = sum(1)
           ,@c2 = sum(case when x.keyno is null then 0 else 1 end)
    from    INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS c
    join    INFORMATION_SCHEMA.KEY_COLUMN_USAGE u1
    on      u1.CONSTRAINT_CATALOG = c.CONSTRAINT_CATALOG
    and     u1.CONSTRAINT_SCHEMA = c.CONSTRAINT_SCHEMA
    and     u1.CONSTRAINT_NAME = c.CONSTRAINT_NAME
    join    INFORMATION_SCHEMA.KEY_COLUMN_USAGE u2
    on      u2.CONSTRAINT_CATALOG = c.UNIQUE_CONSTRAINT_CATALOG
    and     u2.CONSTRAINT_SCHEMA = c.UNIQUE_CONSTRAINT_SCHEMA
    and     u2.CONSTRAINT_NAME = c.UNIQUE_CONSTRAINT_NAME
    and     u2.ORDINAL_POSITION = u1.ORDINAL_POSITION
    join
    (
        select  1 keyno, 'ObjectName' lkey, 'ObjectName' fkey
        union select  2, 'Field', 'FieldName'
    ) x
    on      x.keyno = u1.ORDINAL_POSITION
    and     x.lkey = u1.COLUMN_NAME
    and     x.fkey = u2.COLUMN_NAME
    where   c.CONSTRAINT_NAME = 'shlParametersFieldFK'
    and     u1.TABLE_NAME = 'shlParameters'
    and     u2.TABLE_NAME = 'shlFields'

    if @c1 <> @c2 or @c1 <>  2
    begin
        print 'changing foreign key ''shlParametersFieldFK'''
        alter table dbo.shlParameters drop constraint shlParametersFieldFK
    end
end

if object_id('shlParametersFieldFK') is null
begin
    print 'creating foreign key ''shlParametersFieldFK'''
    alter table dbo.shlParameters add constraint shlParametersFieldFK
    foreign key (ObjectName,Field) references dbo.shlFields(ObjectName,FieldName)
end
go

print '.oOo.'
go

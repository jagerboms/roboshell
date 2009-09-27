print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlFields') is null
begin
    create table dbo.shlFields
    (
        ObjectName varchar(32) not null
       ,FieldName varchar(32) not null
       ,Sequence integer not null
       ,Label varchar(100) null
       ,Width integer not null
       ,DisplayType varchar(3) not null -- (T)ext, (L)abel, (D)ropdown list, (C)heck, (H)idden ...
       ,FillProcess varchar(32) null
       ,TextField varchar(32) null
       ,ValueField varchar(200) null
       ,DisplayWidth integer not null
       ,Format varchar(50) null
       ,IsPrimary char(1) not null
       ,Justify char(1) not null     -- (L)eft, (R)ight, (C)enter or (D)efault
       ,Enabled char(1) not null
       ,Required char(1) not null
       ,Locate char(1) not null      -- (N)ormal, new (C)olumn, new (G)roup
       ,ValueType varchar(25) not null
       ,HelpText varchar(200) null
       ,LabelWidth integer not null
       ,Decimals integer null
       ,NullText varchar(200) null
       ,DisplayHeight integer null
       ,LinkColumn varchar(32) null
       ,LinkField varchar(32) null
       ,Container varchar(32) null
       ,constraint shlFieldsPK primary key clustered
       (
            ObjectName
           ,FieldName
       )
    )
end
go

if not exists
(
    select  'a'
    from INFORMATION_SCHEMA.Columns c
    where   c.TABLE_NAME = 'shlFields'
    and     c.TABLE_SCHEMA = 'dbo'
    and     c.COLUMN_NAME = 'Container'
)
begin
    alter table dbo.shlFields
        add Container varchar(32) null

    alter table dbo.shlFields
        alter column DisplayType varchar(3) not null
end
go

print '.oOo.'
go

print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlParameters') is null
begin
    create table dbo.shlParameters
    (
        ObjectName varchar(32) not null
       ,ParameterName varchar(32) not null
       ,Sequence integer not null
       ,IsInput char(1) not null
       ,IsOutput char(1) not null
       ,ValueType varchar(25) not null
       ,Width integer not null
       ,Value varchar(255) null
       ,Type char(1) not null
       ,Field varchar(32) null
       ,constraint shlParametersPK primary key clustered
       (
            ObjectName
           ,ParameterName
       )
    )
end
go

if not exists
(
    select  'a'
    from INFORMATION_SCHEMA.Columns c
    where   c.TABLE_NAME = 'shlParameters'
    and     c.TABLE_SCHEMA = 'dbo'
    and     c.COLUMN_NAME = 'Field'
)
begin
    alter table dbo.shlParameters
        add Field varchar(32) null
end
go

print '.oOo.'
go

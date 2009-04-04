print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlActions') is null
begin
    create table dbo.shlActions
    (
        ObjectName varchar(32) not null
       ,ActionName varchar(32) not null
       ,Sequence integer not null
       ,Process varchar(32) null
       ,Enabled char(1) not null
       ,RowBased char(1) not null
       ,Validate char(1) not null
       ,CloseObject char(1) not null
       ,IsDblClick char(1) not null
       ,IsButton char(1) not null
       ,ImageFile varchar(128) null
       ,ToolTip varchar(128) null
       ,MenuType char(1) not null
       ,MenuText varchar(40) null
       ,Parent varchar(32) null
       ,IsKey char(1) not null
       ,KeyCode integer null
       ,Shift varchar(12) null
       ,FieldName varchar(32) null
       ,ProcessField varchar(32) null
       ,LinkedParam varchar(32) null
       ,ParamValue varchar(200) null
       ,constraint shlActionsPK primary key clustered
       (
            ObjectName
           ,ActionName
       )
    )
end
go

if not exists
(
    select  'a'
    from INFORMATION_SCHEMA.Columns c
    where   c.TABLE_NAME = 'shlActions'
    and     c.TABLE_SCHEMA = 'dbo'
    and     c.COLUMN_NAME = 'LinkedParam'
)
begin
    alter table dbo.shlActions add
        LinkedParam varchar(32) null
       ,ParamValue varchar(200) null
end
go

print '.oOo.'
go

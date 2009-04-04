print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlProperties') is null
begin
    create table dbo.shlProperties
    (
        ObjectName varchar(32) not null
       ,PropertyType char(2) not null
       ,PropertyName varchar(32) not null
       ,Value varchar(200) not null
       ,constraint shlPropertiesPK primary key clustered
       (
            ObjectName
           ,PropertyType
           ,PropertyName
       )
    )
end
go

print '.oOo.'
go

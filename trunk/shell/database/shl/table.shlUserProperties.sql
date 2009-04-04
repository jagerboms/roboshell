print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlUserProperties') is null
begin
    create table dbo.shlUserProperties
    (
        ObjectName varchar(32) not null
       ,PropertyName varchar(32) not null
       ,UserName sysname not null
       ,Value varchar(2000) not null
       ,constraint shlUserPropertiesPK primary key clustered
       (
            UserName
           ,ObjectName
           ,PropertyName
       )
    )
end
go

print '.oOo.'
go

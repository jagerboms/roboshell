print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlObjects') is null
begin
    create table dbo.shlObjects
    (
        ObjectName varchar(32) not null
       ,ObjectType varchar(32) not null
       ,constraint shlObjectsPK primary key clustered
       (
            ObjectName
       )
    )
end
go

print '.oOo.'
go

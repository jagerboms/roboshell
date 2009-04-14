print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

if object_id('dbo.helpColours') is null
begin
    create table dbo.helpColours
    (
        SystemID varchar(12) not null
       ,ObjectName varchar(32) not null
       ,ColourValue varchar(200) not null
       ,ValueDescription varchar(30) null
       ,State varchar(2) not null
       ,AuditID integer not null
       ,constraint helpColoursPK primary key clustered
       (
            SystemID
           ,ObjectName
           ,ColourValue
       )
    )
end
go

print '.oOo.'
go

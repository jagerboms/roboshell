print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

if object_id('dbo.helpObjects') is null
begin
    create table dbo.helpObjects
    (
        SystemID varchar(12) not null
       ,ObjectName varchar(32) not null
       ,HelpText varchar(4000) null
       ,ColourText varchar(2000) null
       ,State varchar(2) not null
       ,AuditID integer not null
       ,constraint helpObjectsPK primary key clustered
       (
            SystemID
           ,ObjectName
       )
    )
end
go

print '.oOo.'
go

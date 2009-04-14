print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

if object_id('dbo.helpSystems') is null
begin
    create table dbo.helpSystems
    (
        SystemID varchar(12) not null
       ,SystemName varchar(100) null
       ,Copyright varchar(100) not null
       ,State varchar(2) not null
       ,AuditID integer not null
       ,constraint helpSystemsPK primary key clustered
       (
            SystemID
       )
    )
end
go

print '.oOo.'
go

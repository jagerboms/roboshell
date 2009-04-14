print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

if object_id('dbo.helpActions') is null
begin
    create table dbo.helpActions
    (
        SystemID varchar(12) not null
       ,ObjectName varchar(32) not null
       ,ActionName varchar(32) not null
       ,HelpText varchar(4000) null
       ,State varchar(2) not null
       ,AuditID integer not null
       ,constraint helpActionsPK primary key clustered
       (
            SystemID
           ,ObjectName
           ,ActionName
       )
    )
end
go

print '.oOo.'
go

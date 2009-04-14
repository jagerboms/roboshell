print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

if object_id('dbo.helpFields') is null
begin
    create table dbo.helpFields
    (
        SystemID varchar(12) not null
       ,ObjectName varchar(32) not null
       ,FieldName varchar(32) not null
       ,HelpText varchar(4000) null
       ,State varchar(2) not null
       ,AuditID integer not null
       ,constraint helpFieldsPK primary key clustered
       (
            SystemID
           ,ObjectName
           ,FieldName
       )
    )
end
go

print '.oOo.'
go

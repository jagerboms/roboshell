print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlModuleProcesses') is null
begin
    create table dbo.shlModuleProcesses
    (
        ModuleID    varchar(32) not null
       ,ProcessName varchar(32) not null
       ,constraint shlModuleProcessesPK primary key clustered
        (
            ModuleID
           ,ProcessName
        )
    )
end
go

print '.oOo.'
go

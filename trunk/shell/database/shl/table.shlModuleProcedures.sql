print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlModuleProcedures') is null
begin
    create table dbo.shlModuleProcedures
    (
        ModuleID      varchar(32) not null
       ,ProcedureName varchar(32) not null
       ,constraint shlModuleProceduresPK primary key clustered
        (
            ModuleID
           ,ProcedureName
        )
    )
end
go

print '.oOo.'
go

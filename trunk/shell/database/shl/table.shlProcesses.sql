print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlProcesses') is null
begin
    create table dbo.shlProcesses
    (
        ProcessName varchar(32) not null
       ,SuccessProcess varchar(32) null
       ,FailProcess varchar(32) null
       ,ConfirmMsg varchar(100) null
       ,UpdateParent char(1) not null
       ,ObjectName varchar(32) not null
       ,dbo char(1) not null
       ,LoadVariables char(1) not null
       ,constraint shlProcessesPK primary key clustered
       (
           ProcessName
       )
    )
end
go

print '.oOo.'
go

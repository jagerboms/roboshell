print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if not exists (select 'a' from master.dbo.sysmessages where error = 51000)
    execute master.dbo.sp_addmessage 51000, 16, 'This %s already exists.'
go
if not exists (select 'a' from master.dbo.sysmessages where error = 51001)
    execute master.dbo.sp_addmessage 51001, 16, '%s not found.'
go

if not exists (select 'a' from master.dbo.sysmessages where error = 51002)
    execute master.dbo.sp_addmessage 51002, 16, 'Nothing has changed.'
go

print '.oOo.'
go

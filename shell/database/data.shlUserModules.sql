print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if not exists
(
    select  'a'
    from    dbo.shlModuleUsers m
    where   m.UserName = 'db_owner'
)
begin
    insert into dbo.shlModuleUsers (ModuleID, UserName, GrantDeny)
    values ('base', 'db_owner', 'G')
    print 'db_owner permission set.'
end
go

print '.oOo.'
go


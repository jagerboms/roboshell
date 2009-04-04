print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlRoleMembersGet') is not null
begin
    drop procedure dbo.shlRoleMembersGet
end
go

create Procedure dbo.shlRoleMembersGet
    @UserName sysname
as
begin
    set nocount on
    execute sp_helprolemember
        @rolename = @Username
end
go

print '.oOo.'
go

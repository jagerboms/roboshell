print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

if object_id('dbo.helpSystemModules') is not null
begin
    drop function dbo.helpSystemModules
end
go
create function dbo.helpSystemModules()
returns @rs table
(
    ObjectName varchar(32) not null
)
as
begin
    insert into @rs
    select  'shlSecurityAdmin'
    union
    select  'shlUserPerm'
    union
    select  'shlGroupMember'
    union
    select  'shlRoleMembers'

    return
end
go

print '.oOo.'
go

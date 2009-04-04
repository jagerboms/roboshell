print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if object_id('dbo.shlSecurityAdminGet') is not null
begin
    drop procedure dbo.shlSecurityAdminGet
end
go

create Procedure dbo.shlSecurityAdminGet
as
begin
    set nocount on

    select  UserName = s.name
           ,LoginName = case when suser_sname(s.sid) is null then null else
                    dbo.shlUserNameGet(suser_sname(s.sid)) end
           ,Role = case s.issqlrole when 1 then 'R' else
	               case s.isntgroup when 1 then 'G' else
        	           case s.isntuser when 1 then 'U' else 'S'
                   end end end
           ,Type = case s.issqlrole when 1 then 'Role' else
                       case s.isntgroup when 1 then 'Group' else
                           case s.isntuser when 1 then 'User' else 'SQL'
                   end end end
           ,Permissioned = case coalesce(u.PermCount, 0) when 0 then 'No' else 'Yes' end
    from    dbo.sysusers s
    left join
    (
        select  m.UserName
               ,PermCount = count(*)
        from    dbo.shlModuleUsers m
        group by m.UserName
    ) u
    on      u.UserName = s.name
    where   name not in ( 'db_owner', 'db_accessadmin', 'db_securityadmin', 'db_ddladmin',
                'db_backupoperator', 'db_datareader', 'db_datawriter', 'db_denydatareader',
                'db_denydatawriter', 'INFORMATION_SCHEMA', 'sys' )
end
go

print '.oOo.'
go

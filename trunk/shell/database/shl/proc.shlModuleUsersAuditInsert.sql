print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if object_id('dbo.shlModuleUsersAuditInsert') is not null
begin
    drop procedure dbo.shlModuleUsersAuditInsert
end
go

create procedure dbo.shlModuleUsersAuditInsert
    @ModuleID varchar(32)
   ,@UserName sysname
   ,@ActionType char(1)
as
begin
    set nocount on
    declare @e integer
           ,@AuditID integer

    set @e = 0
    while @e = 0
    begin
        select  @AuditID = max(a.AuditID)
        from    dbo.shlModuleUsersAudit a
        where   a.ModuleID = @ModuleID
        and     a.UserName = @UserName
    
        insert into dbo.shlModuleUsersAudit
        (
            ModuleID, UserName, AuditID,
            GrantDeny, ActionType, AuditTime, UserID
        )
        select  @ModuleID
               ,@UserName
               ,coalesce(@AuditID, 0) + 1
               ,mu.GrantDeny
               ,@ActionType
               ,getdate()
               ,suser_sname()
        from    dbo.shlModuleUsers mu
        where   mu.ModuleID = @ModuleID
        and     mu.UserName = @UserName

        set @e = @@error
        break
    end
    return @e
end
go

print '.oOo.'
go

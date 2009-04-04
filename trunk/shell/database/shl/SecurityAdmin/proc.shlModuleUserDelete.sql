print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go
if object_id('dbo.shlModuleUserDelete') is not null
begin
    drop procedure dbo.shlModuleUserDelete
end
go

create Procedure dbo.shlModuleUserDelete
    @ModuleID varchar(32)
   ,@UserName sysname
as
begin
    set nocount on
    declare @e integer

    set @e = 0
    while @e = 0
    begin
        begin transaction

-- write to audit trail...

        execute @e = dbo.shlModuleUsersAuditInsert
            @ModuleID = @ModuleID
           ,@UserName = @UserName
           ,@ActionType = 'D'
        if @e <> 0
        begin
            break
        end

        delete
        from    dbo.shlModuleUsers
        where   ModuleID = @ModuleID
        and     UserName = @UserName

        set @e = @@error
        if @e <> 0
        begin
            break
        end

        break
    end
    if @e <> 0
    begin
        if @@trancount > 0
        begin
            rollback transaction
        end
    end
    else
    begin
        if @@trancount > 0
        begin
            commit transaction
        end
    end
    execute dbo.shlGrantRevoke
end
go

print '.oOo.'
go

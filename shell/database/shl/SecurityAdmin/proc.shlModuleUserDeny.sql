print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go
if object_id('dbo.shlModuleUserDeny') is not null
begin
    drop procedure dbo.shlModuleUserDeny
end
go

create Procedure dbo.shlModuleUserDeny
    @ModuleID varchar(32)
   ,@UserName sysname
as
begin
    set nocount on
    declare @e integer

    set @e = 0
    while @e = 0
    begin

-- validate @UserName

        if not exists
        (
            select  'a'
            from    dbo.sysusers u
            where   u.name = @UserName
        )
        begin
            set @e = 50010
            raiserror (@e,-1,-1,'User is not valid in this database')
            break
        end

-- validate @ModuleID/@ModuleAction

        if not exists
        (
            select  'a'
            from    dbo.shlModules m
            where   m.ModuleID = @ModuleID
        )
        begin
            set @e = 50011
                        raiserror (@e,-1,-1,'Module does not exist')
            break
        end

        begin transaction

        if exists
        (
            select  'a'
            from    dbo.shlModuleUsers m
            where   m.ModuleID = @ModuleID
            and     m.UserName = @UserName
        )
        begin
            update  dbo.shlModuleUsers
            set     GrantDeny = 'D'
            where   ModuleID = @ModuleID
            and     UserName = @UserName

            set @e = @@error
            if @e <> 0
            begin
                break
            end

-- write to audit trail...

            execute @e = dbo.shlModuleUsersAuditInsert
                @ModuleID = @ModuleID
               ,@UserName = @UserName
               ,@ActionType = 'U'
            if @e <> 0
            begin
                break
            end
        end
        else
        begin
            insert into dbo.shlModuleUsers
            (
                ModuleID, UserName, GrantDeny
            )
            values
            (
                @ModuleID, @UserName, 'D'
            )
            set @e = @@error
            if @e <> 0
            begin
                break
            end

-- write to audit trail...

            execute @e = dbo.shlModuleUsersAuditInsert
                @ModuleID = @ModuleID
               ,@UserName = @UserName
               ,@ActionType = 'I'
            if @e <> 0
            begin
                break
            end
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

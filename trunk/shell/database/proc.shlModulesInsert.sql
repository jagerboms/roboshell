print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if object_id('dbo.shlModulesInsert') is not null
begin
    drop procedure dbo.shlModulesInsert
end
go

create Procedure dbo.shlModulesInsert
    @ModuleID varchar(32)
   ,@OwnerModule varchar(32)
   ,@Description varchar(50) = null
as
begin
    set nocount on
    declare @e integer

    set @e = 0
    while @e = 0
    begin
        begin transaction

        update  dbo.shlModules
        set     Description = coalesce(@Description, @ModuleID)
        where   ModuleID = @ModuleID

        if @@rowcount = 0 
        begin
            insert into dbo.shlModules
            (
                ModuleID, Description
            )
            values
            (
                @ModuleID, coalesce(@Description, @ModuleID)
            )
            set @e = @@error
            if @e <> 0
            begin
                break
            end

            insert into dbo.shlModuleOwners
            (
                ModuleID, OwnerModule
            )
            values
            (
                @ModuleID, @OwnerModule
            )
            set @e = @@error
            if @e <> 0
            begin
                break
            end
        end
        else
        begin
            update  dbo.shlModuleOwners
            set     OwnerModule = @OwnerModule
            where   ModuleID = @ModuleID

            set @e = @@error
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
end
go

print '.oOo.'
go

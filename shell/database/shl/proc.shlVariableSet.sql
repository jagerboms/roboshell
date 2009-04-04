print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if object_id('dbo.shlVariableSet') is not null
begin
    drop procedure dbo.shlVariableSet
end
go

create Procedure dbo.shlVariableSet
    @VariableID varchar(32)
   ,@VariableValue varchar(200)
   ,@ShellUse char(1) = null
as
begin
    set nocount on

    declare @e integer
           ,@AuditID integer
           ,@tran integer

    set @e = 0
    while @e = 0
    begin
        set @tran = @@trancount
        if @tran = 0        -- do not start a new transaction
        begin
            begin transaction
        end

        select  @AuditID = v.AuditID
        from    dbo.shlVariables v (holdlock)
        where   v.VariableID = @VariableID

        if @@rowcount = 0
        begin
            insert into dbo.shlVariables
            (
                VariableID, VariableValue, ShellUse, AuditID
            )
            values
            (
                @VariableID, @VariableValue, coalesce(@Shelluse, 'N'), 1
            )
            set  @e = @@error
            if @e <> 0
            begin
                break
            end

            execute @e = dbo.shlVariablesAuditInsert
                @VariableID = @VariableID
               ,@AuditID = 1
               ,@ActionType = 'I'
        end
        else
        begin
            set @AuditID = @AuditID + 1

            update  dbo.shlVariables
            set     VariableValue = @VariableValue
                   ,ShellUse = coalesce(@ShellUse, ShellUse)
                   ,AuditID = @AuditID
            where   VariableID = @VariableID
    
            set  @e = @@error
            if @e <> 0
            begin
                break
            end

            execute @e = dbo.shlVariablesAuditInsert
                @VariableID = @VariableID
               ,@AuditID = @AuditID
               ,@ActionType = 'U'
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
        if @@trancount > 0 and @tran = 0
        begin
            commit transaction    -- do not commit if transaction was initiated 
        end                       -- outside this procedure
    end
end
go

print '.oOo.'
go

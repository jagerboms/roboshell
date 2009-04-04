print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if object_id('dbo.shlVariablesAuditInsert') is not null
begin
    drop procedure dbo.shlVariablesAuditInsert
end
go

create procedure dbo.shlVariablesAuditInsert
    @VariableID varchar(32)
   ,@AuditID   integer
   ,@ActionType char(1)
as
begin
    set nocount on
    declare @e integer

    set @e = 0
    while @e = 0
    begin
        insert into dbo.shlVariablesAudit
        (
            VariableID, AuditID,
            VariableValue, ShellUse, ActionType,
            AuditTime, UserID
        )
        select  @VariableID
               ,@AuditID
               ,k.VariableValue
               ,k.ShellUse
               ,@ActionType
               ,getdate()
               ,suser_sname()
        from    dbo.shlVariables k
        where   k.VariableID = @VariableID

        set @e = @@error
        break
    end
    return @e
end
go

print '.oOo.'
go

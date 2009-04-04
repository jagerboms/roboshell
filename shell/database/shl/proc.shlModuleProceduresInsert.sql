print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlModuleProceduresInsert') is not null
begin
    drop procedure dbo.shlModuleProceduresInsert
end
go

create procedure dbo.shlModuleProceduresInsert
    @ProcedureName varchar(32)
   ,@ModuleID varchar(32)
as
begin
    set nocount on
    declare @e integer

    set @e = 0
    while @e = 0
    begin
        print rtrim(@ModuleID) + ': ' + @ProcedureName

        if object_id(@ProcedureName) is null
        begin
            set @e = 50040
            raiserror @e 'Invalid procedure name'
            break
        end

        if not exists
        (
            select  'a'
            from    dbo.shlModuleProcedures p
            where   p.ModuleID = @ModuleID
            and     p.ProcedureName = @ProcedureName
        )
        begin
            insert into dbo.shlModuleProcedures
            (
                ModuleID, ProcedureName
            )
            values
            (
                @ModuleID, @ProcedureName
            )
            set @e = @@error
            if @e <> 0
            begin
                break
            end
        end
        break
    end
    return @e
end
go

print '.oOo.'
go

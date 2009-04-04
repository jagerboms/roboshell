print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlProcessesDelete') is not null
begin
    drop procedure dbo.shlProcessesDelete
end
go

create procedure dbo.shlProcessesDelete
    @ProcessName varchar(32)
as
begin
    set nocount on
    declare @e integer

    set @e = 0
    while @e = 0
    begin

        begin transaction

        delete
        from    dbo.shlModuleProcesses
        where   ProcessName = @ProcessName

        set @e = @@error
        if @e <> 0
        begin
            break
        end

        delete
        from    dbo.shlProcesses
        where   ProcessName = @ProcessName

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
    return @e
end
go

print '.oOo.'
go

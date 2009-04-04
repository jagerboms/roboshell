print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlModuleProcessInsert') is not null
begin
    drop procedure dbo.shlModuleProcessInsert
end
go

create procedure dbo.shlModuleProcessInsert
    @ProcessName varchar(32)
   ,@ModuleID varchar(32)
as
begin
    set nocount on
    declare @e integer

    set @e = 0
    while @e = 0
    begin
        print rtrim(@ModuleID) + ': ' + @ProcessName

        if not exists
        (
            select  'a'
            from    dbo.shlModuleProcesses p
            where   p.ModuleID = @ModuleID
            and     p.ProcessName = @ProcessName
        )
        begin
            insert into dbo.shlModuleProcesses
            (
                ModuleID, ProcessName
            )
            values
            (
                @ModuleID, @ProcessName
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

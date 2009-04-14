print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlProcessesInsert') is not null
begin
    drop procedure dbo.shlProcessesInsert
end
go

create procedure dbo.shlProcessesInsert
    @ProcessName varchar(32)
   ,@ModuleID varchar(32) = null
   ,@ObjectName varchar(32)
   ,@SuccessProcess varchar(32) = null
   ,@FailProcess varchar(32) = null
   ,@ConfirmMsg varchar(100) = null
   ,@UpdateParent char(1) = 'N'     -- (Y)es, (S)uspend or (N)o
   ,@dbo char(1) = 'N'
   ,@LoadVariables char(1) = 'N'
as
begin
    set nocount on
    declare @e integer
           ,@m varchar(200)
           ,@c integer

    set @e = 0
    while @e = 0
    begin
        print 'Process: ' + @ProcessName

        if @ModuleID is not null
        begin
            if not exists
            (
                select  'a'
                from    dbo.shlModules m
                where   m.ModuleID = @ModuleID
            )
            begin
                set @e = 60600
                set @m = 'Error module ' + @ModuleID + ' does not exist...'
                raiserror @e @m
                break
            end
        end

        if @ObjectName is not null
        begin
            if not exists
            (
                select  'a'
                from    dbo.shlObjects o
                where   o.ObjectName = @ObjectName
            )
            begin
                set @e = 60600
                set @m = 'Error object ' + @ObjectName + ' does not exist...'
                raiserror @e @m
                break
            end
        end

        if @SuccessProcess is not null
        begin
            if @SuccessProcess = @ProcessName
            begin
                set @e = 60600
                set @m = 'Error success process ' + @SuccessProcess + ' can not be itself...'
                raiserror @e @m
                break
            end
            if not exists
            (
                select  'a'
                from    dbo.shlProcesses p
                where   p.ProcessName = @SuccessProcess
            )
            begin
                set @e = 60600
                set @m = 'Error success process ' + @SuccessProcess + ' does not exist...'
                raiserror @e @m
                break
            end
        end

        if @FailProcess is not null
        begin
            if @SuccessProcess = @ProcessName
            begin
                set @e = 60600
                set @m = 'Error exit fail process ' + @FailProcess + ' can not be itself...'
                raiserror @e @m
                break
            end
            if not exists
            (
                select  'a'
                from    dbo.shlProcesses p
                where   p.ProcessName = @FailProcess
            )
            begin
                set @e = 60600
                set @m = 'Error fail process ' + @FailProcess + ' does not exist...'
                raiserror @e @m
                break
            end
        end

        set @UpdateParent = upper(@UpdateParent)
        if @UpdateParent not in ('Y', 'S')
        begin
            set @UpdateParent = 'N'
        end

        set @dbo = upper(@dbo)
        if @dbo <> 'Y'
        begin
            set @dbo = 'N'
        end

        set @LoadVariables = upper(@LoadVariables)
        if @LoadVariables <> 'Y'
        begin
            set @LoadVariables = 'N'
        end

        begin transaction

        delete
        from    dbo.shlModuleProcesses
        where   ProcessName = @ProcessName

        set @e = @@error
        if @e <> 0
        begin
            break
        end

        update  dbo.shlProcesses
        set     SuccessProcess = @SuccessProcess
               ,FailProcess = @FailProcess
               ,ConfirmMsg = @ConfirmMsg
               ,UpdateParent = @UpdateParent
               ,ObjectName = @ObjectName
               ,dbo = @dbo
               ,LoadVariables = @LoadVariables
        where   ProcessName = @ProcessName

        select  @e = @@error
               ,@c = @@rowcount
        if @e <> 0
        begin
            break
        end

        if @c = 0
        begin
            insert into dbo.shlProcesses
            (
                ProcessName, SuccessProcess, FailProcess,
                ConfirmMsg, UpdateParent, ObjectName,
                dbo, LoadVariables
            )
            values
            (
                @ProcessName, @SuccessProcess, @FailProcess,
                @ConfirmMsg, @UpdateParent, @ObjectName,
                @dbo, @LoadVariables
            )
            set @e = @@error
            if @e <> 0
            begin
                break
            end
        end

        if @ModuleID is not null
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

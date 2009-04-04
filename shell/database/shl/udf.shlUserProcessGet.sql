print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlUserProcessGet') is not null
begin
    drop function dbo.shlUserProcessGet
end
go

create function dbo.shlUserProcessGet()
returns @rs table
(
    ProcessName varchar(32) collate database_default not null
   ,SuccessProcess varchar(32) collate database_default null
   ,FailProcess varchar(32) collate database_default null
   ,ConfirmMsg varchar(100) collate database_default null
   ,UpdateParent char(1) collate database_default not null
   ,ObjectName varchar(32) collate database_default not null
   ,LoadVariables char(1) collate database_default not null
)
as
begin
    declare @err integer
           ,@count integer
           ,@dbo char(1)

    set @err = 0
    while @err = 0
    begin
        if coalesce(is_member('db_owner'), 0) = 1
        or coalesce(is_member('db_securityadmin'), 0) = 1
        begin
            set @dbo = 'Y'
        end
        else
        begin
            set @dbo = 'N'
        end

        insert into @rs
        select  distinct
                p.ProcessName
               ,p.SuccessProcess
               ,p.FailProcess
               ,p.ConfirmMsg
               ,p.UpdateParent
               ,p.ObjectName
               ,p.LoadVariables
        from    dbo.shlProcesses p
        join    dbo.shlModuleProcesses mp
        on      p.ProcessName = mp.ProcessName
        join    dbo.shlModulePermGet() n
        on      n.ModuleID = mp.ModuleID
        and     n.GrantDeny = 'G'
        where   p.dbo = 'N' or @dbo = 'Y'

        select  @err = @@error
               ,@count = @@rowcount
        if @err <> 0
        begin
            break
        end

        insert into @rs
        select  distinct
                p.ProcessName
               ,p.SuccessProcess
               ,p.FailProcess
               ,p.ConfirmMsg
               ,p.UpdateParent
               ,p.ObjectName
               ,p.LoadVariables
        from    @rs r
        join    dbo.shlFields f
        on      f.ObjectName = r.ObjectName
        join    dbo.shlProcesses p
        on      f.FillProcess = p.ProcessName
        and    (p.dbo = 'N' or @dbo = 'Y')
        where   not exists
        (
            select  'a'
            from    @rs s
            where   s.ProcessName = p.ProcessName
        )

        set @err = @@error
        if @err <> 0
        begin
            break
        end

        while @err = 0 and @count <> 0
        begin
            delete
            from    @rs
            where   coalesce(SuccessProcess, '') <> ''
            and     not exists
            (
                select  'a'
                from    @rs n1
                join    dbo.shlProcesses p
                on      n1.ProcessName = p.ProcessName
                and    (p.dbo = 'N' or @dbo = 'Y')
                left join @rs n2
                on      p.SuccessProcess = n2.ProcessName
            )
            select  @err = @@error
                   ,@count = @@rowcount
            if @err <> 0
            begin
                break
            end

            delete
            from    @rs
            where   coalesce(FailProcess, '') <> ''
            and not exists
            (
                select  'a'
                from    @rs n1
                join    dbo.shlProcesses p
                on      n1.ProcessName = p.ProcessName
                and    (p.dbo = 'N' or @dbo = 'Y')
                join    @rs n2
                on      p.FailProcess = n2.ProcessName
            )
            select  @err = @@error
                   ,@count = @count + @@rowcount
            if @err <> 0
            begin
                break
            end
        end
        break
    end
    return
end
go

print '.oOo.'
go

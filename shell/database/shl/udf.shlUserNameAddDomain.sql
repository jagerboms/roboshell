print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if object_id('dbo.shlUserNameAddDomain') is not null
begin
    drop function dbo.shlUserNameAddDomain
end
go

create function dbo.shlUserNameAddDomain
(
    @Name varchar(200) = null
)
returns varchar(200)
as
begin
    declare @u varchar(200)
           ,@i integer

    set @u = coalesce(@Name, suser_sname())
    if charindex('\', @Name, 0) < 1
    begin
        if not exists
        (
            select  'a'
            from    dbo.sysusers u
            where   u.name = @Name
            and     u.isntgroup <> 1
            and     u.isntuser <> 1
        )
        begin
            set @u = suser_sname()
            set @i = charindex('\', @u, 0)
            set @u = substring(@u, 1, @i) + @Name
        end
    end
    return (@u)
end
go

print '.oOo.'
go


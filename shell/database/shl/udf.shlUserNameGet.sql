print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlUserNameGet') is not null
begin
    drop function dbo.shlUserNameGet
end
go

create function dbo.shlUserNameGet
(
    @Name sysname = null
)
returns sysname
as
begin
    declare @u sysname
           ,@i integer

    set @u = coalesce(@Name, suser_sname())
    set @i = charindex('\', @u, 0)
    if @i > 0
    begin
        set @u = substring(@u, @i + 1, 200)
    end
    return (@u)
end
go

print '.oOo.'
go

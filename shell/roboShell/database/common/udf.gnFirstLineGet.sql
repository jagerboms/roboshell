print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

if object_id('dbo.gnFirstLineGet') is not null
begin
    drop function dbo.gnFirstLineGet
end
go

create function dbo.gnFirstLineGet(@Text varchar(4000))
returns varchar(80)
as
begin
    declare
        @i integer
       ,@s varchar(80)

    set @i = charindex(char(13), @Text)
    set @i = @i - 1
    if @i > 80 set @i = 80
    if @i <= 0 set @i = 80
    if @i > len(@Text) set @i = len(@Text)

    if @i > 0
    begin
        set @s = substring(@Text, 1, @i)
    end
    else
    begin
        set @s = null
    end
    return (@s)
end
go

print '.oOo.'
go

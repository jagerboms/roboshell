print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlHTMLEncode') is not null
begin
    drop function dbo.shlHTMLEncode
end
go

create function dbo.shlHTMLEncode(@Text varchar(8000))
returns varchar(8000)
as
begin
    declare @html varchar(8000)
    set @html = replace(@Text, '&', '&amp;')
    set @html = replace(@html, '>', '&gt;')
    set @html = replace(@html, '<', '&lt;')
    set @html = replace(@html, '"', '&quot;')
    set @html = replace(@html, char(13) + char(10), '<br>')
    set @html = replace(@html, char(13), '<br>')
    set @html = replace(@html, char(10), '<br>')
    return (@html)
end
go

print '.oOo.'
go

print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if not exists (select 'a' from dbo.shlStyles where StyleID = 'default')
begin
    execute dbo.shlStylesInsert
        @StyleID = 'default'
       ,@RowForeColor = 'black'
       ,@RowBackColor = 'white'
       ,@SelForeColor = 'black'
       ,@SelBackColor = 'gainsboro'
end
go

if not exists (select 'a' from dbo.shlStyles where StyleID = 'disabled')
begin
    execute dbo.shlStylesInsert
        @StyleID = 'disabled'
       ,@RowForeColor = 'red'
       ,@RowBackColor = 'mistyrose'
       ,@SelForeColor = 'red'
       ,@SelBackColor = 'lightcoral'
end
go

print '.oOo.'
go

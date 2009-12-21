print '-----------------------'
print '-- Akuna Care - Pets --'
print '-----------------------'
set nocount on
go

if object_id('dbo.shlStyles') is null
begin
    print 'creating dbo.shlStyles'
    create table dbo.shlStyles
    (
        StyleID varchar(32) not null
       ,RowForeColor varchar(32) null
       ,RowBackColor varchar(32) null
       ,SelForeColor varchar(32) null
       ,SelBackColor varchar(32) null
       ,constraint shlStylesPK primary key clustered
        (
            StyleID
        )
    )
end
go

print '.oOo.'
go

print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if object_id('dbo.shlTableValues') is null
begin
    print 'creating dbo.shlTableValues'
    create table dbo.shlTableValues
    (
        TableName sysname not null
       ,ColumnName sysname not null
       ,ColumnValue varchar(32) not null
       ,ValueDescription varchar(50) not null
       ,Keys varchar(30) null
       ,constraint shlTableValuesPK primary key clustered
       (
            TableName
           ,ColumnName
           ,ColumnValue
       )
    )
end
go

print '.oOo.'
go

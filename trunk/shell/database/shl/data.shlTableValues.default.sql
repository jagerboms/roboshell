print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print '-----------------------------------'
set nocount on
go
delete from dbo.shlTableValues where TableName = 'default'
go

insert into dbo.shlTableValues (TableName, ColumnName, ColumnValue, ValueDescription)
values('default', 'State', 'ac', 'Active')
go
insert into dbo.shlTableValues (TableName, ColumnName, ColumnValue, ValueDescription)
values('default', 'State', 'dl', 'Disabled')
go
insert into dbo.shlTableValues (TableName, ColumnName, ColumnValue, ValueDescription)
values('default', 'State', 'pd', 'Pending')
go
insert into dbo.shlTableValues (TableName, ColumnName, ColumnValue, ValueDescription)
values('default', 'State', 'cp', 'Complete')
go


insert into dbo.shlTableValues (TableName, ColumnName, ColumnValue, ValueDescription)
values('default', 'ActionType', 'I', 'New')
go
insert into dbo.shlTableValues (TableName, ColumnName, ColumnValue, ValueDescription)
values('default', 'ActionType', 'U', 'Edit')
go
insert into dbo.shlTableValues (TableName, ColumnName, ColumnValue, ValueDescription)
values('default', 'ActionType', 'D', 'Disable')
go

print '.oOo.'
go

print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlValidations') is null
begin
    create table dbo.shlValidations
    (
        ObjectName varchar(32) not null
       ,ValidationName varchar(32) not null
       ,FieldName varchar(32) not null
       ,ValidationType char(2) not null   -- EQ,NE,NN,GT,GE,LT,LE
       ,ValueType char(1) not null        -- Constant, Field, Process
       ,Process varchar(32) null
       ,Value varchar(200) not null
       ,Message varchar(200) not null
       ,ReturnParameter varchar(32) null
       ,constraint shlValidationsPK primary key clustered
       (
            ObjectName
           ,ValidationName
       )
    )
end
go

print '.oOo.'
go

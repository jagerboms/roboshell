print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlPropertiesInsert') is not null
begin
    drop procedure dbo.shlPropertiesInsert
end
go

create procedure dbo.shlPropertiesInsert
    @ObjectName varchar(32)
   ,@PropertyType char(2) = 'df'
   ,@PropertyName varchar(32)
   ,@Value varchar(200)
   ,@Update char(1) = 'N'
as
begin
    set nocount on
    declare @e integer
           ,@c integer

    set @e = 0
    while @e = 0
    begin
        print 'Property: ' + rtrim(@ObjectName) + '.' +  rtrim(@PropertyType) + '.' + @PropertyName

        if upper(@Update) = 'Y'
        begin
            update  dbo.shlProperties
            set     Value = @Value
            where   ObjectName = @ObjectName
            and     PropertyType = @PropertyType
            and     PropertyName = @PropertyName

            select  @e = @@error
                   ,@c = @@rowcount
            if @e <> 0 or @c > 0
            begin
                break
            end
        end

        insert into dbo.shlProperties
        (
            ObjectName, PropertyType, PropertyName, Value
        )
        values
        (
            @ObjectName, @PropertyType, @PropertyName, @Value
        )
        set @e = @@error
        if @e <> 0
        begin
            break
        end
        break
    end
    return @e
end
go

print '.oOo.'
go

print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlUserPropertyAlter') is not null
begin
    drop procedure dbo.shlUserPropertyAlter
end
go

create procedure dbo.shlUserPropertyAlter
    @ObjectName char(32)
   ,@PropertyType char(2) = 'df'
   ,@PropertyName char(32)
   ,@Value varchar(2000) = null
as
begin
    set nocount on
    declare @e integer
           ,@c integer

    set @e = 0
    while @e = 0
    begin
        if coalesce(@Value, '') = ''
        begin
            delete
            from    dbo.shlUserProperties
            where   ObjectName = @ObjectName
            and     PropertyName = @PropertyName
            and     UserName = suser_sname()
    
            set @e = @@error        
            if @e <> 0
            begin
                break
            end
        end
        else
        begin
            update  dbo.shlUserProperties
            set     Value = @Value
            where   ObjectName = @ObjectName
            and     PropertyName = @PropertyName
            and     UserName = suser_sname()
    
            select  @e = @@error        
                   ,@c = @@rowcount
            if @e <> 0
            begin
                break
            end

            if @c = 0
            begin
                insert into dbo.shlUserProperties
                (
                    ObjectName, PropertyName,
                    UserName, Value
                )
                values
                (
                    @ObjectName, @PropertyName,
                    suser_sname(), @Value
                )
                set @e = @@error
                if @e <> 0
                begin
                    break
                end
            end
        end
        break
    end
    return @e
end
go

print '.oOo.'
go

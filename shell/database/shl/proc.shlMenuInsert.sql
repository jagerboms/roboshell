print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlMenuInsert') is not null
begin
    drop procedure dbo.shlMenuInsert
end
go

create procedure dbo.shlMenuInsert
    @ObjectName varchar(32)
   ,@Update char(1) = 'N'
as
begin
    set nocount on
    declare @e integer

    set @e = 0
    while @e = 0
    begin
        if @Update = 'Y'
        begin
-- save actions set in other modules...
            select  *
            into    #temp
            from    dbo.shlActions
            where   ObjectName = @ObjectName
        end

        execute @e = dbo.shlObjectsInsert
            @ObjectName = @ObjectName
           ,@ObjectType = 'Menu'
        if @e <> 0
        begin
            break
        end

        if @Update = 'Y'
        begin
-- retrieve saved actions
            insert into dbo.shlActions
            select  *
            from    #temp
            
            drop table #temp
        end
        break
    end
    return @e
end
go

print '.oOo.'
go

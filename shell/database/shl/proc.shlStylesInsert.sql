print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlStylesInsert') is not null
begin
    drop procedure dbo.shlStylesInsert
end
go

create procedure dbo.shlStylesInsert
    @StyleID varchar(32)
   ,@RowForeColor varchar(32) = null
   ,@RowBackColor varchar(32) = null
   ,@SelForeColor varchar(32) = null
   ,@SelBackColor varchar(32) = null
as
begin
    set nocount on
    declare @e integer
           ,@m varchar(200)
           ,@c integer
           ,@sequence integer

    set @e = 0
    while @e = 0
    begin
        print 'Style: ' + rtrim(@StyleID)

        update  dbo.shlStyles
        set     RowForeColor = @RowForeColor
               ,RowBackColor = @RowBackColor
               ,SelForeColor = @SelForeColor
               ,SelBackColor = @SelBackColor
        where   StyleID = @StyleID

        select  @e = @@error
               ,@c = @@rowcount
        if @e <> 0 or @c > 0
        begin
            break
        end

        insert into dbo.shlStyles
        (
            StyleID, RowForeColor, RowBackColor,
            SelForeColor, SelBackColor
        )
        values
        (
            @StyleID, @RowForeColor, @RowBackColor,
            @SelForeColor, @SelBackColor
        )
        set @e = @@error
        break
    end
    return @e
end
go

print '.oOo.'
go

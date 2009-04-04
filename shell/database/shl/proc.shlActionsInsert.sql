print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlActionsInsert') is not null
begin
    drop procedure dbo.shlActionsInsert
end
go

create procedure dbo.shlActionsInsert
    @ObjectName varchar(32)
   ,@ActionName varchar(32)
   ,@Process varchar(32) = null
   ,@Enabled char(1) = 'Y'
   ,@RowBased char(1) = 'N'
   ,@Validate char(1) = 'N'
   ,@CloseObject char(1) = 'N'
   ,@IsDblClick char(1) = 'N'
   ,@ImageFile varchar(128) = null
   ,@ToolTip varchar(128) = null
   ,@MenuType char(1) = 'N'
   ,@MenuText varchar(40) = null
   ,@Parent varchar(32) = null
   ,@KeyCode integer = null
   ,@Shift varchar(12) = null
   ,@FieldName varchar(32) = null
   ,@ProcessField varchar(32) = null
   ,@LinkedParam varchar(32) = null     -- parameter linked to the button checked state
   ,@ParamValue varchar(200) = null     -- the || delimited values for parameter for each state
   ,@Update char(1) = 'N'
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
        print 'Action: ' + rtrim(@ObjectName) + '.' + @ActionName 

        if @Process is not null
        begin
            if not exists
            (
                select  'a'
                from    dbo.shlProcesses p
                where   p.ProcessName = @Process
            )
            begin
                set @e = 60600
                set @m = 'Error process ' + @Process + ' does not exist...'
                raiserror @e @m
                break
            end
        end

        set @Enabled = upper(@Enabled)
        if @Enabled <> 'Y'
        begin
            set @Enabled = 'N'
        end

        set @RowBased = upper(@RowBased)
        if @RowBased <> 'Y'
        begin
            set @RowBased = 'N'
        end

        set @Validate = upper(@Validate)
        if @Validate <> 'Y'
        begin
            set @Validate = 'N'
        end

        set @CloseObject = upper(@CloseObject)
        if @CloseObject not in ('Y', 'O', 'P', 'Q')
        begin
            set @CloseObject = 'N'
        end

        set @IsDblClick = upper(@IsDblClick)
        if @IsDblClick <> 'Y'
        begin
            set @IsDblClick = 'N'
        end

        if upper(@Update) = 'Y'
        begin
            delete
            from    dbo.shlActionRules
            where   ObjectName = @ObjectName
            and     ActionName = @ActionName

            delete
            from    dbo.shlActionProcessRules
            where   ObjectName = @ObjectName
            and     ActionName = @ActionName

            update  dbo.shlActions
            set     Process = @Process
                   ,Enabled = @Enabled
                   ,RowBased = @RowBased
                   ,Validate = @Validate
                   ,CloseObject = @CloseObject
                   ,IsDblClick = @IsDblClick
                   ,IsButton = case when coalesce(@ImageFile, '') = '' then 'N' else 'Y' end
                   ,ImageFile = @ImageFile
                   ,ToolTip = @ToolTip
                   ,MenuType = @MenuType
                   ,MenuText = @MenuText
                   ,Parent = @Parent
                   ,IsKey = case when coalesce(@KeyCode, 0) < 1 then 'N' else 'Y' end
                   ,KeyCode = @KeyCode
                   ,Shift = @Shift
                   ,FieldName = @FieldName
                   ,ProcessField = @ProcessField
                   ,LinkedParam = @LinkedParam
                   ,ParamValue = @ParamValue
            where   ObjectName = @ObjectName
            and     ActionName = @ActionName

            select  @e = @@error
                   ,@c = @@rowcount
            if @e <> 0 or @c > 0
            begin
                break
            end
        end

        select  @sequence = max(Sequence)
        from    dbo.shlActions p
        where   ObjectName = @ObjectName

        set @sequence = coalesce(@sequence, 0) + 1

        insert into dbo.shlActions
        (
            ObjectName, ActionName, Sequence, Process,
            Enabled, RowBased, Validate, CloseObject,
            IsDblClick, IsButton, ImageFile, ToolTip,
            MenuType, MenuText, Parent, IsKey, 
            KeyCode, Shift, FieldName, ProcessField,
            LinkedParam, ParamValue
        )
        select  @ObjectName
               ,@ActionName
               ,@Sequence
               ,@Process
               ,@Enabled
               ,@RowBased
               ,@Validate
               ,@CloseObject
               ,@IsDblClick
               ,case when coalesce(@ImageFile, '') = '' then 'N' else 'Y' end
               ,@ImageFile
               ,@ToolTip
               ,@MenuType
               ,@MenuText
               ,@Parent
               ,case when coalesce(@KeyCode, 0) < 1 then 'N' else 'Y' end
               ,@KeyCode
               ,@Shift
               ,@FieldName
               ,@ProcessField
               ,@LinkedParam
               ,@ParamValue
        set @e = @@error
        break
    end
    return @e
end
go

print '.oOo.'
go

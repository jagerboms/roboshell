print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlFieldsInsert') is not null
begin
    drop procedure dbo.shlFieldsInsert
end
go

create procedure dbo.shlFieldsInsert
    @ObjectName varchar(32)            -- identity of object owning field/column
   ,@FieldName varchar(32)             -- Unique field identifier
   ,@Label varchar(100) = null         -- dialog label / grid column header 
   ,@Width integer = 0                 -- for valuetype 'string' the storage width
   ,@DisplayWidth integer              -- display item width / default grid column width
   ,@DisplayHeight integer = 1         -- number of lines to display (text and label display types on dialog only)
   ,@ValueType varchar(25)             -- 'string' 'integer' 'double' 'currency' 'object' etc
   ,@DisplayType varchar(3) = 'T'      -- (H)idden, d(R)opdown field, (L)abel, (B)ordered label, (T)extbox, (D)ropdown list, li(S)tbox, (C)heckbox
   ,@FillProcess varchar(32) = null    -- shell process to retrieve data for dropdown or listbox controls
   ,@TextField varchar(32) = null      -- column of datatable to display
   ,@ValueField varchar(200) = null    -- column of datatable to use as field value 
   ,@LinkColumn varchar(32) = null     -- datatable column to filter on
   ,@LinkField varchar(32) = null      -- field providing filter value
   ,@Format varchar(50) = null         -- display format string
   ,@IsPrimary char(1) = 'N'           -- indicates if field is part of unique row identification
   ,@Justify char(1) = 'D'             -- (D)efault, (R)ight, (L)eft or (C)enter
   ,@Enabled char(1) = 'Y'             -- Is the control editable?
   ,@Required char(1) = 'N'            -- Is the control mandatory?
   ,@Locate char(1) = 'N'              -- (N)ormal, (C)olumn, (G)roup or (P)air
   ,@HelpText varchar(200) = null      -- text displayed when entering data into field
   ,@LabelWidth integer = 100          -- width of field label (dialog)
   ,@Decimals integer = -1             -- Number of decimal places to display
   ,@NullText varchar(200) = null      -- Text to display when raw data is null value
   ,@Container varchar(32) = null      -- 
as
begin
    set nocount on
    declare @e integer
           ,@sequence integer

    set @e = 0
    while @e = 0
    begin
        print 'Field: ' + rtrim(@ObjectName) + '.' + @FieldName 

        set @IsPrimary = upper(@IsPrimary)
        if @IsPrimary <> 'Y'
        begin
            set @IsPrimary = 'N'
        end

        set @Enabled = upper(@Enabled)
        if @Enabled <> 'Y'
        begin
            set @Enabled = 'N'
        end

        set @Required = upper(@Required)
        if @Required <> 'Y'
        begin
            set @Required = 'N'
        end

        set @Locate = upper(@Locate)
        if @Locate not in ('C', 'G', 'P')
        begin
            set @Locate = 'N'
        end

        select  @sequence = max(Sequence)
        from    dbo.shlFields p
        where   ObjectName = @ObjectName

        set @sequence = coalesce(@sequence, 0) + 1

        if @sequence = 1 and @Locate <> 'N'
        begin
            set @e = 50055
            raiserror @e 'You cannot set field locate on first field'
            break
        end

        insert into dbo.shlFields
        (
            ObjectName, FieldName, Sequence, Label,
            Width, DisplayType, FillProcess, TextField,
            ValueField, DisplayWidth, DisplayHeight, Format,
            IsPrimary, Justify, Enabled, Required,
            Locate, ValueType, HelpText, LabelWidth,
            Decimals, NullText,
            LinkColumn, LinkField, Container
        )
        values
        (
            @ObjectName, @FieldName, @Sequence, coalesce(@Label, @FieldName),
            @Width, @DisplayType, @FillProcess, @TextField,
            @ValueField, @DisplayWidth, @DisplayHeight, @Format,
            @IsPrimary, @Justify, @Enabled, @Required,
            @Locate, @ValueType, @HelpText, @LabelWidth,
            @Decimals, @NullText,
            @LinkColumn, @LinkField, @Container
        )
        set @e = @@error
        break
    end
    return @e
end
go

print '.oOo.'
go

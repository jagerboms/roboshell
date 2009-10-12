print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlFieldParamInsert') is not null
begin
    drop procedure dbo.shlFieldParamInsert
end
go

create procedure dbo.shlFieldParamInsert
    @ObjectName varchar(32)          -- identity of object owning field/column
   ,@FieldName varchar(32)           -- Unique field identifier
   ,@Label varchar(100) = null       -- dialog label / grid column header 
   ,@Width integer = 0               -- for valuetype 'string' the storage width
   ,@DisplayWidth integer            -- display item width / default grid column width
   ,@DisplayHeight integer = 1       -- number of lines to display (text and label display types on dialog only)
   ,@ValueType varchar(25)           -- 'string' 'integer' 'double' 'currency' 'object' etc
   ,@DisplayType varchar(3) = 'T'    -- (H)idden, d(R)opdown field, (L)abel, (B)ordered label, (T)extbox, (D)ropdown list, li(S)tbox, (C)heckbox
   ,@FillProcess varchar(32) = null  -- shell process to retrieve data for dropdown or listbox controls
   ,@TextField varchar(32) = null    -- column of datatable to display
   ,@ValueField varchar(200) = null  -- column of datatable to use as field value 
                                     -- or '||' delimited data elements.
   ,@LinkColumn varchar(32) = null   -- datatable column to filter on
   ,@LinkField varchar(32) = null    -- field providing filter value
   ,@Format varchar(50) = null       -- display format string
   ,@IsPrimary char(1) = 'N'         -- indicates if field is part of unique row identification
   ,@Justify char(1) = 'D'           -- (D)efault, (R)ight, (L)eft or (C)enter
   ,@Enabled char(1) = 'Y'           -- Is the control editable?
   ,@Required char(1) = 'N'          -- Is the control mandatory?
   ,@Locate char(1) = 'N'            -- (N)ormal, (C)olumn, (G)roup or (P)air
   ,@HelpText varchar(200) = null    -- text displayed when entering data into field
   ,@LabelWidth integer = 100        -- width of field label (dialog)
   ,@Decimals integer = -1           -- Number of decimal places to display
   ,@NullText varchar(200) = null    -- Text to display when raw data is null value

   ,@ParamValue varchar(200) = null  -- Initial parameter value
   ,@IsInput char(1) = 'Y'           -- Is this an input parameter?
   ,@IsOutput char(1) = 'Y'          -- Is this an output parameter?
   ,@ParamType char(1) = 'U'         -- (U)ser or (C)onectionstring
   ,@Container varchar(32) = null    -- 
as
begin
    set nocount on
    declare @e integer
           ,@sequence integer
           ,@Cap varchar(100)

    set @e = 0
    while @e = 0
    begin
        set @Cap = coalesce(@Label, @FieldName)

        begin transaction

        execute @e = dbo.shlFieldsInsert
            @ObjectName = @ObjectName
           ,@FieldName = @FieldName
           ,@Label = @Cap
           ,@Width = @Width
           ,@DisplayWidth = @DisplayWidth
           ,@DisplayHeight = @DisplayHeight
           ,@ValueType = @ValueType
           ,@DisplayType = @DisplayType
           ,@FillProcess = @FillProcess
           ,@TextField = @TextField
           ,@ValueField = @ValueField
           ,@LinkColumn = @LinkColumn
           ,@LinkField = @LinkField
           ,@Format = @Format
           ,@IsPrimary = @IsPrimary
           ,@Justify = @Justify
           ,@Enabled = @Enabled
           ,@Required = @Required
           ,@Locate = @Locate
           ,@HelpText = @HelpText
           ,@LabelWidth = @LabelWidth
           ,@Decimals = @Decimals
           ,@NullText = @NullText
           ,@Container = @Container
        if @e <> 0
        begin
            break
        end

        execute @e = dbo.shlParametersInsert
            @ObjectName = @ObjectName
           ,@ParameterName = @FieldName
           ,@ValueType = @ValueType
           ,@Width = @Width
           ,@Value = @ParamValue
           ,@IsInput = @IsInput
           ,@IsOutput = @IsOutput
           ,@Type = @ParamType
           ,@Field = @FieldName
        if @e <> 0
        begin
            break
        end

        break
    end
    if @e <> 0
    begin
        if @@trancount > 0
        begin
            rollback transaction
        end
    end
    else
    begin
        if @@trancount > 0
        begin
            commit transaction
        end
    end
    return @e
end
go

print '.oOo.'
go

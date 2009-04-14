print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlGridFormInsert
    @ObjectName = 'helpColours'
   ,@Title = 'Colours'
   ,@TitleParameters = 'pSystemID'
   ,@DataParameter = 'helpColoursGet'
   ,@StateFilter = 'Y'
   ,@ColourColumn = 'State'
   ,@HelpPage = 'helpColours.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpColoursx'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpColours'
go

execute dbo.shlProcessesInsert
    @ProcessName = 'helpColours'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'SystemIDTrans'
   ,@SuccessProcess = 'helpColoursx'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpColours'
   ,@ParameterName = 'pSystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpColours'
   ,@FieldName = 'AuditID'
   ,@DisplayWidth = -1
   ,@ValueType = 'integer'
   ,@DisplayType = 'H'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpColours'
   ,@FieldName = 'SystemID'
   ,@Label = 'System'
   ,@ValueType = 'string'
   ,@Width = 12
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
   ,@IsPrimary = 'Y'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpColours'
   ,@FieldName = 'ObjectName'
   ,@Label = 'Object'
   ,@ValueType = 'string'
   ,@Width = 32
   ,@DisplayWidth = 80
   ,@IsPrimary = 'Y'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpColours'
   ,@FieldName = 'ColourValue'
   ,@Label = 'Value'
   ,@ValueType = 'string'
   ,@Width = 200
   ,@DisplayWidth = 150
   ,@IsPrimary = 'Y'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpColours'
   ,@FieldName = 'ValueDescription'
   ,@Label = 'Description'
   ,@ValueType = 'string'
   ,@Width = 30
   ,@DisplayWidth = 80
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpColours'
   ,@FieldName = 'StateName'
   ,@Label = 'State'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@DisplayWidth = 50
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpColours'
   ,@FieldName = 'State'
   ,@ValueType = 'string'
   ,@Width = 2
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpColours'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'dl'
   ,@Value = 'red'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpColours'
   ,@PropertyType = 'cb'
   ,@PropertyName = 'dl'
   ,@Value = 'mistyrose'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpColours'
   ,@PropertyType = 'cb'
   ,@PropertyName = 'sh'
   ,@Value = 'yellowgreen'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpColours'
   ,@PropertyType = 'cb'
   ,@PropertyName = 'ft'
   ,@Value = 'aliceblue'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpColours'
   ,@PropertyType = 'lk'
   ,@PropertyName = 'HELPCOLOUR'
   ,@Value = ''
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'helpColours'
   ,@ActionName = 'Refresh'
   ,@Process = 'helpColoursGet'
   ,@ImageFile = 'refresh.gif'
   ,@ToolTip = 'Refresh data'
   ,@KeyCode = 120
go

execute dbo.shlActionsInsert
    @ObjectName = 'helpColours'
   ,@ActionName = 'Update'
   ,@Process = 'helpColoursEdit'
   ,@RowBased = 'Y'
   ,@Validate = 'Y'
   ,@ImageFile = 'edit.gif'
   ,@ToolTip = 'Amend Colour'
go

execute dbo.shlActionsInsert
    @ObjectName = 'helpColours'
   ,@ActionName = 'History'
   ,@Process = 'helpColoursAudit'
   ,@RowBased = 'Y'
   ,@ImageFile = 'history.gif'
   ,@ToolTip = 'Change history'
go

print '.oOo.'
go

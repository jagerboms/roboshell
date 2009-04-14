print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlGridFormInsert
    @ObjectName = 'helpActions'
   ,@Title = 'Actions'
   ,@TitleParameters = 'pSystemID'
   ,@DataParameter = 'helpActionsGet'
   ,@StateFilter = 'Y'
   ,@ColourColumn = 'State'
   ,@HelpPage = 'helpActions.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpActionsx'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpActions'
go

execute dbo.shlProcessesInsert
    @ProcessName = 'helpActions'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'SystemIDTrans'
   ,@SuccessProcess = 'helpActionsx'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpActions'
   ,@ParameterName = 'pSystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpActions'
   ,@FieldName = 'AuditID'
   ,@DisplayWidth = -1
   ,@ValueType = 'integer'
   ,@DisplayType = 'H'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpActions'
   ,@FieldName = 'SystemID'
   ,@Label = 'System'
   ,@ValueType = 'string'
   ,@Width = 12
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
   ,@IsPrimary = 'Y'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpActions'
   ,@FieldName = 'ObjectName'
   ,@Label = 'Object'
   ,@ValueType = 'string'
   ,@Width = 32
   ,@DisplayWidth = 80
   ,@IsPrimary = 'Y'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpActions'
   ,@FieldName = 'ActionName'
   ,@Label = 'Action'
   ,@ValueType = 'string'
   ,@Width = 32
   ,@DisplayWidth = 80
   ,@IsPrimary = 'Y'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpActions'
   ,@FieldName = 'Description'
   ,@ValueType = 'string'
   ,@Width = 80
   ,@DisplayWidth = 200
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpActions'
   ,@FieldName = 'HelpText'
   ,@Label = 'Text'
   ,@ValueType = 'string'
   ,@Width = 4000
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpActions'
   ,@FieldName = 'StateName'
   ,@Label = 'State'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@DisplayWidth = 50
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpActions'
   ,@FieldName = 'State'
   ,@ValueType = 'string'
   ,@Width = 2
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpActions'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'dl'
   ,@Value = 'red'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpActions'
   ,@PropertyType = 'cb'
   ,@PropertyName = 'dl'
   ,@Value = 'mistyrose'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpActions'
   ,@PropertyType = 'cb'
   ,@PropertyName = 'sh'
   ,@Value = 'yellowgreen'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpActions'
   ,@PropertyType = 'cb'
   ,@PropertyName = 'ft'
   ,@Value = 'aliceblue'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpActions'
   ,@PropertyType = 'lk'
   ,@PropertyName = 'HELPACTION'
   ,@Value = ''
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'helpActions'
   ,@ActionName = 'Refresh'
   ,@Process = 'helpActionsGet'
   ,@ImageFile = 'refresh.gif'
   ,@ToolTip = 'Refresh data'
   ,@KeyCode = 120
go

execute dbo.shlActionsInsert
    @ObjectName = 'helpActions'
   ,@ActionName = 'Update'
   ,@Process = 'helpActionsEdit'
   ,@RowBased = 'Y'
   ,@Validate = 'Y'
   ,@ImageFile = 'edit.gif'
   ,@ToolTip = 'Amend Action'
go

execute dbo.shlActionsInsert
    @ObjectName = 'helpActions'
   ,@ActionName = 'History'
   ,@Process = 'helpActionsAudit'
   ,@RowBased = 'Y'
   ,@ImageFile = 'history.gif'
   ,@ToolTip = 'Change history'
go

print '.oOo.'
go

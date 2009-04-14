print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlGridFormInsert
    @ObjectName = 'helpFields'
   ,@Title = 'Fields'
   ,@TitleParameters = 'pSystemID'
   ,@DataParameter = 'helpFieldsGet'
   ,@StateFilter = 'Y'
   ,@ColourColumn = 'State'
   ,@HelpPage = 'helpFields.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpFieldsx'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpFields'
go

execute dbo.shlProcessesInsert
    @ProcessName = 'helpFields'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'SystemIDTrans'
   ,@SuccessProcess = 'helpFieldsx'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpFields'
   ,@ParameterName = 'pSystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpFields'
   ,@FieldName = 'AuditID'
   ,@DisplayWidth = -1
   ,@ValueType = 'integer'
   ,@DisplayType = 'H'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpFields'
   ,@FieldName = 'SystemID'
   ,@Label = 'System'
   ,@ValueType = 'string'
   ,@Width = 12
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
   ,@IsPrimary = 'Y'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpFields'
   ,@FieldName = 'ObjectName'
   ,@Label = 'Object'
   ,@ValueType = 'string'
   ,@Width = 32
   ,@DisplayWidth = 80
   ,@IsPrimary = 'Y'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpFields'
   ,@FieldName = 'FieldName'
   ,@Label = 'Field'
   ,@ValueType = 'string'
   ,@Width = 32
   ,@DisplayWidth = 80
   ,@IsPrimary = 'Y'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpFields'
   ,@FieldName = 'Description'
   ,@ValueType = 'string'
   ,@Width = 80
   ,@DisplayWidth = 200
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpFields'
   ,@FieldName = 'HelpText'
   ,@Label = 'Text'
   ,@ValueType = 'string'
   ,@Width = 4000
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpFields'
   ,@FieldName = 'StateName'
   ,@Label = 'State'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@DisplayWidth = 50
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpFields'
   ,@FieldName = 'State'
   ,@ValueType = 'string'
   ,@Width = 2
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpFields'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'dl'
   ,@Value = 'red'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpFields'
   ,@PropertyType = 'cb'
   ,@PropertyName = 'dl'
   ,@Value = 'mistyrose'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpFields'
   ,@PropertyType = 'cb'
   ,@PropertyName = 'sh'
   ,@Value = 'yellowgreen'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpFields'
   ,@PropertyType = 'cb'
   ,@PropertyName = 'ft'
   ,@Value = 'aliceblue'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpFields'
   ,@PropertyType = 'lk'
   ,@PropertyName = 'HELPFIELD'
   ,@Value = ''
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'helpFields'
   ,@ActionName = 'Refresh'
   ,@Process = 'helpFieldsGet'
   ,@ImageFile = 'refresh.gif'
   ,@ToolTip = 'Refresh data'
   ,@KeyCode = 120
go

execute dbo.shlActionsInsert
    @ObjectName = 'helpFields'
   ,@ActionName = 'Update'
   ,@Process = 'helpFieldsEdit'
   ,@RowBased = 'Y'
   ,@Validate = 'Y'
   ,@ImageFile = 'edit.gif'
   ,@ToolTip = 'Amend Field'
go

execute dbo.shlActionsInsert
    @ObjectName = 'helpFields'
   ,@ActionName = 'History'
   ,@Process = 'helpFieldsAudit'
   ,@RowBased = 'Y'
   ,@ImageFile = 'history.gif'
   ,@ToolTip = 'Change history'
go

print '.oOo.'
go

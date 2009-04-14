print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlGridFormInsert
    @ObjectName = 'helpFieldsAudit'
   ,@Title = 'Field Change History'
   ,@DataParameter = 'helpFieldsAuditGet'
   ,@ColourColumn = 'ActionType'
   ,@TitleParameters = 'SystemID||ObjectName||FieldName'
   ,@HelpPage = 'helpFieldsAudit.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpFieldsAudit'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpFieldsAudit'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsAudit'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsAudit'
   ,@ParameterName = 'ObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsAudit'
   ,@ParameterName = 'FieldName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpFieldsAudit'
   ,@FieldName = 'AuditID'
   ,@Label = 'Sequence'
   ,@ValueType = 'integer'
   ,@DisplayWidth = 40
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpFieldsAudit'
   ,@FieldName = 'ActionType'
   ,@ValueType = 'string'
   ,@Width = 1
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpFieldsAudit'
   ,@FieldName = 'Action'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@DisplayWidth = 60
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpFieldsAudit'
   ,@FieldName = 'UserID'
   ,@Label = 'Actioned By'
   ,@ValueType = 'string'
   ,@Width = 128
   ,@DisplayWidth = 100
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpFieldsAudit'
   ,@FieldName = 'AuditTime'
   ,@Label = 'Actioned'
   ,@ValueType = 'datetime'
   ,@DisplayWidth = 80
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpFieldsAudit'
   ,@FieldName = 'Description'
   ,@ValueType = 'string'
   ,@Width = 80
   ,@DisplayWidth = 200
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpFieldsAudit'
   ,@FieldName = 'State'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@DisplayWidth = 50
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpFieldsAudit'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'I'
   ,@Value = 'black'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpFieldsAudit'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'D'
   ,@Value = 'red'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpFieldsAudit'
   ,@PropertyType = 'cb'
   ,@PropertyName = 'D'
   ,@Value = 'mistyrose'
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'helpFieldsAudit'
   ,@ActionName = 'Refresh'
   ,@Process = 'helpFieldsAuditGet'
   ,@ImageFile = 'refresh.gif'
   ,@ToolTip = 'Refresh data'
   ,@KeyCode = 120
go

print '.oOo.'
go

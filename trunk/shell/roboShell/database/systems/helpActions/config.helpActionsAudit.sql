print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlGridFormInsert
    @ObjectName = 'helpActionsAudit'
   ,@Title = 'Action Change History'
   ,@DataParameter = 'helpActionsAuditGet'
   ,@ColourColumn = 'ActionType'
   ,@TitleParameters = 'SystemID||ObjectName||ActionName'
   ,@HelpPage = 'helpActionsAudit.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpActionsAudit'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpActionsAudit'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsAudit'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsAudit'
   ,@ParameterName = 'ObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsAudit'
   ,@ParameterName = 'ActionName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpActionsAudit'
   ,@FieldName = 'AuditID'
   ,@Label = 'Sequence'
   ,@ValueType = 'integer'
   ,@DisplayWidth = 40
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpActionsAudit'
   ,@FieldName = 'ActionType'
   ,@ValueType = 'string'
   ,@Width = 1
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpActionsAudit'
   ,@FieldName = 'Action'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@DisplayWidth = 60
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpActionsAudit'
   ,@FieldName = 'UserID'
   ,@Label = 'Actioned By'
   ,@ValueType = 'string'
   ,@Width = 128
   ,@DisplayWidth = 100
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpActionsAudit'
   ,@FieldName = 'AuditTime'
   ,@Label = 'Actioned'
   ,@ValueType = 'datetime'
   ,@DisplayWidth = 80
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpActionsAudit'
   ,@FieldName = 'Description'
   ,@ValueType = 'string'
   ,@Width = 80
   ,@DisplayWidth = 200
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpActionsAudit'
   ,@FieldName = 'State'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@DisplayWidth = 50
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpActionsAudit'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'I'
   ,@Value = 'black'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpActionsAudit'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'D'
   ,@Value = 'red'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpActionsAudit'
   ,@PropertyType = 'cb'
   ,@PropertyName = 'D'
   ,@Value = 'mistyrose'
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'helpActionsAudit'
   ,@ActionName = 'Refresh'
   ,@Process = 'helpActionsAuditGet'
   ,@ImageFile = 'refresh.gif'
   ,@ToolTip = 'Refresh data'
   ,@KeyCode = 120
go

print '.oOo.'
go

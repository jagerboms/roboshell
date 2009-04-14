print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlGridFormInsert
    @ObjectName = 'SystemsAudit'
   ,@Title = 'System Change History'
   ,@DataParameter = 'SystemsAuditGet'
   ,@ColourColumn = 'ActionType'
   ,@TitleParameters = 'SystemID'
   ,@HelpPage = 'SystemsAudit.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'SystemsAudit'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'SystemsAudit'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsAudit'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'SystemsAudit'
   ,@FieldName = 'AuditID'
   ,@Label = 'Sequence'
   ,@ValueType = 'integer'
   ,@DisplayWidth = 40
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'SystemsAudit'
   ,@FieldName = 'ActionType'
   ,@ValueType = 'string'
   ,@Width = 1
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'SystemsAudit'
   ,@FieldName = 'Action'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@DisplayWidth = 60
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'SystemsAudit'
   ,@FieldName = 'UserID'
   ,@Label = 'Actioned By'
   ,@ValueType = 'string'
   ,@Width = 128
   ,@DisplayWidth = 100
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'SystemsAudit'
   ,@FieldName = 'AuditTime'
   ,@Label = 'Actioned'
   ,@ValueType = 'datetime'
   ,@DisplayWidth = 80
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'SystemsAudit'
   ,@FieldName = 'SystemName'
   ,@ValueType = 'string'
   ,@Width = 100
   ,@DisplayWidth = 500
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'SystemsAudit'
   ,@FieldName = 'Copyright'
   ,@ValueType = 'string'
   ,@Width = 100
   ,@DisplayWidth = 500
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'SystemsAudit'
   ,@FieldName = 'State'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@DisplayWidth = 50
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'SystemsAudit'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'I'
   ,@Value = 'black'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'SystemsAudit'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'D'
   ,@Value = 'red'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'SystemsAudit'
   ,@PropertyType = 'cb'
   ,@PropertyName = 'D'
   ,@Value = 'mistyrose'
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'SystemsAudit'
   ,@ActionName = 'Refresh'
   ,@Process = 'SystemsAuditGet'
   ,@ImageFile = 'refresh.gif'
   ,@ToolTip = 'Refresh data'
   ,@KeyCode = 120
go

print '.oOo.'
go

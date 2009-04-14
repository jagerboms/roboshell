print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlGridFormInsert
    @ObjectName = 'helpColoursAudit'
   ,@Title = 'Colour Change History'
   ,@DataParameter = 'helpColoursAuditGet'
   ,@ColourColumn = 'ActionType'
   ,@TitleParameters = 'SystemID||ObjectName||ColourValue'
   ,@HelpPage = 'helpColoursAudit.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpColoursAudit'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpColoursAudit'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpColoursAudit'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpColoursAudit'
   ,@ParameterName = 'ObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpColoursAudit'
   ,@ParameterName = 'ColourValue'
   ,@ValueType = 'string'
   ,@Width = 200
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpColoursAudit'
   ,@FieldName = 'AuditID'
   ,@Label = 'Sequence'
   ,@ValueType = 'integer'
   ,@DisplayWidth = 40
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpColoursAudit'
   ,@FieldName = 'ActionType'
   ,@ValueType = 'string'
   ,@Width = 1
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpColoursAudit'
   ,@FieldName = 'Action'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@DisplayWidth = 60
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpColoursAudit'
   ,@FieldName = 'UserID'
   ,@Label = 'Actioned By'
   ,@ValueType = 'string'
   ,@Width = 128
   ,@DisplayWidth = 100
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpColoursAudit'
   ,@FieldName = 'AuditTime'
   ,@Label = 'Actioned'
   ,@ValueType = 'datetime'
   ,@DisplayWidth = 80
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpColoursAudit'
   ,@FieldName = 'ValueDescription'
   ,@ValueType = 'string'
   ,@Width = 30
   ,@DisplayWidth = 150
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpColoursAudit'
   ,@FieldName = 'State'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@DisplayWidth = 50
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpColoursAudit'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'I'
   ,@Value = 'black'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpColoursAudit'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'D'
   ,@Value = 'red'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpColoursAudit'
   ,@PropertyType = 'cb'
   ,@PropertyName = 'D'
   ,@Value = 'mistyrose'
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'helpColoursAudit'
   ,@ActionName = 'Refresh'
   ,@Process = 'helpColoursAuditGet'
   ,@ImageFile = 'refresh.gif'
   ,@ToolTip = 'Refresh data'
   ,@KeyCode = 120
go

print '.oOo.'
go

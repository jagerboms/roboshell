print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlGridFormInsert
    @ObjectName = 'helpObjectAudit'
   ,@Title = 'Object Change History'
   ,@DataParameter = 'helpObjectAuditGet'
   ,@ColourColumn = 'ActionType'
   ,@TitleParameters = 'SystemID||ObjectName'
   ,@HelpPage = 'helpObjectAudit.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpObjectAudit'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpObjectAudit'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpObjectAudit'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpObjectAudit'
   ,@ParameterName = 'ObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObjectAudit'
   ,@FieldName = 'AuditID'
   ,@Label = 'Sequence'
   ,@ValueType = 'integer'
   ,@DisplayWidth = 40
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObjectAudit'
   ,@FieldName = 'ActionType'
   ,@ValueType = 'string'
   ,@Width = 1
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObjectAudit'
   ,@FieldName = 'Action'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@DisplayWidth = 60
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObjectAudit'
   ,@FieldName = 'UserID'
   ,@Label = 'Actioned By'
   ,@ValueType = 'string'
   ,@Width = 128
   ,@DisplayWidth = 100
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObjectAudit'
   ,@FieldName = 'AuditTime'
   ,@Label = 'Actioned'
   ,@ValueType = 'datetime'
   ,@DisplayWidth = 80
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObjectAudit'
   ,@FieldName = 'HelpText'
   ,@ValueType = 'string'
   ,@Width = 80
   ,@DisplayWidth = 200
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObjectAudit'
   ,@FieldName = 'ColourText'
   ,@ValueType = 'string'
   ,@Width = 80
   ,@DisplayWidth = 200
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObjectAudit'
   ,@FieldName = 'State'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@DisplayWidth = 50
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpObjectAudit'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'I'
   ,@Value = 'black'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpObjectAudit'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'D'
   ,@Value = 'red'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpObjectAudit'
   ,@PropertyType = 'cb'
   ,@PropertyName = 'D'
   ,@Value = 'mistyrose'
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'helpObjectAudit'
   ,@ActionName = 'Refresh'
   ,@Process = 'helpObjectAuditGet'
   ,@ImageFile = 'refresh.gif'
   ,@ToolTip = 'Refresh data'
   ,@KeyCode = 120
go

print '.oOo.'
go

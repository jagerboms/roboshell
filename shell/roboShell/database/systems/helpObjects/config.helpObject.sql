print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlGridFormInsert
    @ObjectName = 'helpObject'
   ,@Title = 'Objects'
   ,@TitleParameters = 'pSystemID'
   ,@DataParameter = 'helpObjectGet'
   ,@StateFilter = 'Y'
   ,@ColourColumn = 'State'
   ,@HelpPage = 'helpObject.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpObjectx'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpObject'
go

execute dbo.shlProcessesInsert
    @ProcessName = 'helpObject'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'SystemIDTrans'
   ,@SuccessProcess = 'helpObjectx'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpObject'
   ,@ParameterName = 'pSystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObject'
   ,@FieldName = 'AuditID'
   ,@DisplayWidth = -1
   ,@ValueType = 'integer'
   ,@DisplayType = 'H'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObject'
   ,@FieldName = 'SystemID'
   ,@Label = 'System'
   ,@ValueType = 'string'
   ,@Width = 12
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
   ,@IsPrimary = 'Y'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObject'
   ,@FieldName = 'ObjectName'
   ,@Label = 'Object'
   ,@ValueType = 'string'
   ,@Width = 32
   ,@DisplayWidth = 80
   ,@IsPrimary = 'Y'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObject'
   ,@FieldName = 'Description'
   ,@ValueType = 'string'
   ,@Width = 80
   ,@DisplayWidth = 200
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObject'
   ,@FieldName = 'Colour'
   ,@ValueType = 'string'
   ,@Width = 80
   ,@DisplayWidth = 200
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObject'
   ,@FieldName = 'HelpText'
   ,@ValueType = 'string'
   ,@Width = 4000
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObject'
   ,@FieldName = 'ColourText'
   ,@Label = 'Colour'
   ,@ValueType = 'string'
   ,@Width = 2000
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObject'
   ,@FieldName = 'StateName'
   ,@Label = 'State'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@DisplayWidth = 50
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObject'
   ,@FieldName = 'State'
   ,@ValueType = 'string'
   ,@Width = 2
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpObject'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'dl'
   ,@Value = 'red'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpObject'
   ,@PropertyType = 'cb'
   ,@PropertyName = 'dl'
   ,@Value = 'mistyrose'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpObject'
   ,@PropertyType = 'cb'
   ,@PropertyName = 'sh'
   ,@Value = 'yellowgreen'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpObject'
   ,@PropertyType = 'cb'
   ,@PropertyName = 'ft'
   ,@Value = 'aliceblue'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpObject'
   ,@PropertyType = 'lk'
   ,@PropertyName = 'HELPOBJECT'
   ,@Value = ''
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'helpObject'
   ,@ActionName = 'Refresh'
   ,@Process = 'helpObjectGet'
   ,@ImageFile = 'refresh.gif'
   ,@ToolTip = 'Refresh data'
   ,@KeyCode = 120
go

execute dbo.shlActionsInsert
    @ObjectName = 'helpObject'
   ,@ActionName = 'Update'
   ,@Process = 'helpObjectEdit'
   ,@RowBased = 'Y'
   ,@Validate = 'Y'
   ,@ImageFile = 'edit.gif'
   ,@ToolTip = 'Amend Object'
go

execute dbo.shlActionsInsert
    @ObjectName = 'helpObject'
   ,@ActionName = 'History'
   ,@Process = 'helpObjectAudit'
   ,@RowBased = 'Y'
   ,@ImageFile = 'history.gif'
   ,@ToolTip = 'Change history'
go

print '.oOo.'
go

print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlGridFormInsert
    @ObjectName = 'Systems'
   ,@Title = 'Systems'
   ,@DataParameter = 'SystemsGet'
   ,@StateFilter = 'Y'
   ,@ColourColumn = 'State'
   ,@HelpPage = 'Systems.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'Systems'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'Systems'
go

---------------------------------------------------

execute dbo.shlFieldParamInsert
    @ObjectName = 'Systems'
   ,@FieldName = 'AuditID'
   ,@DisplayWidth = -1
   ,@ValueType = 'integer'
   ,@DisplayType = 'H'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'Systems'
   ,@FieldName = 'SystemID'
   ,@Label = 'ID'
   ,@ValueType = 'string'
   ,@Width = 12
   ,@DisplayWidth = 50
   ,@IsPrimary = 'Y'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'Systems'
   ,@FieldName = 'SystemName'
   ,@Label = 'Name'
   ,@ValueType = 'string'
   ,@Width = 100
   ,@DisplayWidth = 150
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'Systems'
   ,@FieldName = 'Copyright'
   ,@ValueType = 'string'
   ,@Width = 100
   ,@DisplayWidth = 150
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'Systems'
   ,@FieldName = 'StateName'
   ,@Label = 'State'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@DisplayWidth = 50
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'Systems'
   ,@FieldName = 'State'
   ,@ValueType = 'string'
   ,@Width = 2
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'Systems'
   ,@PropertyType = 'cl'
   ,@PropertyName = 'dl'
   ,@Value = 'red'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'Systems'
   ,@PropertyType = 'cb'
   ,@PropertyName = 'dl'
   ,@Value = 'mistyrose'
go

execute dbo.shlPropertiesInsert
    @ObjectName = 'Systems'
   ,@PropertyType = 'lk'
   ,@PropertyName = 'SYSTEM'
   ,@Value = ''
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'Systems'
   ,@ActionName = 'Refresh'
   ,@Process = 'SystemsGet'
   ,@ImageFile = 'refresh.gif'
   ,@ToolTip = 'Refresh data'
   ,@KeyCode = 120
go

execute dbo.shlActionsInsert
    @ObjectName = 'Systems'
   ,@ActionName = 'Add'
   ,@Process = 'SystemsAdd'
   ,@ImageFile = 'add.gif'
   ,@ToolTip = 'Create new System'
go

execute dbo.shlActionsInsert
    @ObjectName = 'Systems'
   ,@ActionName = 'Update'
   ,@Process = 'SystemsEdit'
   ,@RowBased = 'Y'
   ,@Validate = 'Y'
   ,@ImageFile = 'edit.gif'
   ,@ToolTip = 'Amend System'

execute dbo.shlActionRulesInsert
    @ObjectName = 'Systems'
   ,@ActionName = 'Update'
   ,@RuleName = 'R1'
   ,@FieldName = 'State'
   ,@Value = 'ac'
go

execute dbo.shlActionsInsert
    @ObjectName = 'Systems'
   ,@ActionName = 'Disable'
   ,@Process = 'SystemsDisable'
   ,@RowBased = 'Y'
   ,@Validate = 'Y'
   ,@ImageFile = 'delete.gif'
   ,@ToolTip = 'Disable System'

execute dbo.shlActionRulesInsert
    @ObjectName = 'Systems'
   ,@ActionName = 'Disable'
   ,@RuleName = 'R1'
   ,@FieldName = 'State'
   ,@Value = 'ac'
go

execute dbo.shlActionsInsert
    @ObjectName = 'Systems'
   ,@ActionName = 'HelpUpdate'
   ,@Process = 'helpToolUpdate'
   ,@RowBased = 'Y'
   ,@ImageFile = 'helpupdate.gif'
   ,@ToolTip = 'Update help tables.'
go

execute dbo.shlActionsInsert
    @ObjectName = 'Systems'
   ,@ActionName = 'HelpBuild'
   ,@Process = 'helpToolbuild'
   ,@RowBased = 'Y'
   ,@ImageFile = 'helpbuild.gif'
   ,@ToolTip = 'Build help pages.'
go

execute dbo.shlActionRulesInsert
    @ObjectName = 'Systems'
   ,@ActionName = 'HelpBuild'
   ,@RuleName = 'R1'
   ,@FieldName = 'SystemID'
   ,@Value = 'DEFAULT'
   ,@ValidationType = 'NE'
go

execute dbo.shlActionsInsert
    @ObjectName = 'Systems'
   ,@ActionName = 'help'
   ,@ImageFile = 'help.gif'
   ,@MenuType = 'S'
   ,@RowBased = 'Y'
go

execute dbo.shlActionsInsert
    @ObjectName = 'Systems'
   ,@ActionName = 'objects'
   ,@Process = 'helpObject'
   ,@MenuType = 'I'
   ,@Parent = 'help'
   ,@MenuText = 'Objects'
   ,@ImageFile = 'objects.gif'
   ,@ToolTip = 'Help page object descriptions'
go

execute dbo.shlActionsInsert
    @ObjectName = 'Systems'
   ,@ActionName = 'fields'
   ,@Process = 'helpFields'
   ,@MenuType = 'I'
   ,@Parent = 'help'
   ,@MenuText = 'Fields'
   ,@ImageFile = 'fields.gif'
   ,@ToolTip = 'Help page field/column descriptions'
go

execute dbo.shlActionsInsert
    @ObjectName = 'Systems'
   ,@ActionName = 'actions'
   ,@Process = 'helpActions'
   ,@MenuType = 'I'
   ,@Parent = 'help'
   ,@MenuText = 'Actions'
   ,@ImageFile = 'action.gif'
   ,@ToolTip = 'Help page action descriptions'
go

execute dbo.shlActionsInsert
    @ObjectName = 'Systems'
   ,@ActionName = 'colours'
   ,@Process = 'helpColours'
   ,@MenuType = 'I'
   ,@Parent = 'help'
   ,@MenuText = 'Colours'
   ,@ImageFile = 'colour.gif'
   ,@ToolTip = 'Help page colour table descriptions'
go

execute dbo.shlActionsInsert
    @ObjectName = 'Systems'
   ,@ActionName = 'History'
   ,@Process = 'SystemsAudit'
   ,@RowBased = 'Y'
   ,@ImageFile = 'history.gif'
   ,@ToolTip = 'Change history'
go

print '.oOo.'
go

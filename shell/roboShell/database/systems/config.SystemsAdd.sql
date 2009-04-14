print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlDialogFormInsert
    @ObjectName = 'SystemsAdd'
   ,@Title = 'Create New System'
   ,@HelpPage = 'SystemsAdd.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'SystemsAdd'
   ,@ModuleID = 'helpsystemmaintain'
   ,@ObjectName = 'SystemsAdd'
   ,@UpdateParent = 'Y'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsAdd'
   ,@ParameterName = 'SystemsGet'
   ,@ValueType = 'object'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'SystemsAdd'
   ,@FieldName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
   ,@DisplayWidth = 50
   ,@IsPrimary = 'Y'
   ,@Required = 'Y'
   ,@IsInput = 'N'
   ,@HelpText = 'help text here'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'SystemsAdd'
   ,@FieldName = 'SystemName'
   ,@ValueType = 'string'
   ,@Width = 100
   ,@DisplayWidth = 150
   ,@IsInput = 'N'
   ,@HelpText = 'help text here'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'SystemsAdd'
   ,@FieldName = 'Copyright'
   ,@ValueType = 'string'
   ,@Width = 100
   ,@DisplayWidth = 150
   ,@Required = 'Y'
   ,@IsInput = 'N'
   ,@HelpText = 'help text here'
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'SystemsAdd'
   ,@ActionName = 'Okay'
   ,@Process = 'SystemsInsert'
   ,@Validate = 'Y'
   ,@CloseObject = 'O'
   ,@ImageFile = 'okay.gif'
   ,@ToolTip = 'Save changes and exit'
   ,@KeyCode = 13
go

execute dbo.shlActionsInsert
    @ObjectName = 'SystemsAdd'
   ,@ActionName = 'Cancel'
   ,@CloseObject = 'Y'
   ,@ImageFile = 'cancel.gif'
   ,@ToolTip = 'Exit without saving changes'
   ,@KeyCode = 27
go

print '.oOo.'
go

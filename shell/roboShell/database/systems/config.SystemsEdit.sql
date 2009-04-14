print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlDialogFormInsert
    @ObjectName = 'SystemsEdit'
   ,@Title = 'Amend System'
   ,@TitleParameters = 'SystemID'
   ,@HelpPage = 'SystemsEdit.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'SystemsEdit'
   ,@ModuleID = 'helpsystemmaintain'
   ,@ObjectName = 'SystemsEdit'
go

---------------------------------------------------

execute dbo.shlFieldParamInsert
    @ObjectName = 'SystemsEdit'
   ,@FieldName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
   ,@DisplayWidth = 50
   ,@DisplayType = 'L'
   ,@IsPrimary = 'Y'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'SystemsEdit'
   ,@FieldName = 'SystemName'
   ,@ValueType = 'string'
   ,@Width = 100
   ,@DisplayWidth = 150
   ,@Enabled = 'Y'
   ,@HelpText = 'help text here'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'SystemsEdit'
   ,@FieldName = 'Copyright'
   ,@ValueType = 'string'
   ,@Width = 100
   ,@DisplayWidth = 150
   ,@Enabled = 'Y'
   ,@Required = 'Y'
   ,@HelpText = 'help text here'
go

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsEdit'
   ,@ParameterName = 'AuditID'
   ,@ValueType = 'integer'
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'SystemsEdit'
   ,@ActionName = 'Okay'
   ,@Process = 'SystemsUpdate'
   ,@RowBased = 'Y'
   ,@Validate = 'Y'
   ,@CloseObject = 'Y'
   ,@ImageFile = 'okay.gif'
   ,@ToolTip = 'Save changes and exit'
   ,@KeyCode = 13
go

execute dbo.shlActionsInsert
    @ObjectName = 'SystemsEdit'
   ,@ActionName = 'Cancel'
   ,@CloseObject = 'Y'
   ,@ImageFile = 'cancel.gif'
   ,@ToolTip = 'Exit without saving changes'
   ,@KeyCode = 27
go

print '.oOo.'
go

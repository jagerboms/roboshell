print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlDialogFormInsert
    @ObjectName = 'helpFieldsEdit'
   ,@Title = 'Amend Field'
   ,@TitleParameters = 'SystemID||ObjectName||FieldName'
   ,@HelpPage = 'helpFieldsEdit.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpFieldsEdit'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpFieldsEdit'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsEdit'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsEdit'
   ,@ParameterName = 'ObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsEdit'
   ,@ParameterName = 'FieldName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpFieldsEdit'
   ,@FieldName = 'HelpText'
   ,@Label = 'Description'
   ,@LabelWidth = 60
   ,@ValueType = 'string'
   ,@Width = 4000
   ,@DisplayWidth = 500
   ,@DisplayHeight = 20
   ,@Enabled = 'Y'
   ,@HelpText = 'Enter field description HTML.'
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsEdit'
   ,@ParameterName = 'AuditID'
   ,@ValueType = 'integer'
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'helpFieldsEdit'
   ,@ActionName = 'Okay'
   ,@Process = 'helpFieldsUpdate'
   ,@RowBased = 'Y'
   ,@Validate = 'Y'
   ,@CloseObject = 'Y'
   ,@ImageFile = 'okay.gif'
   ,@ToolTip = 'Save changes and exit'
   ,@KeyCode = 13
go

execute dbo.shlActionsInsert
    @ObjectName = 'helpFieldsEdit'
   ,@ActionName = 'Cancel'
   ,@CloseObject = 'Y'
   ,@ImageFile = 'cancel.gif'
   ,@ToolTip = 'Exit without saving changes'
   ,@KeyCode = 27
go

print '.oOo.'
go

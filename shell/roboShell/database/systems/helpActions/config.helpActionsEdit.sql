print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlDialogFormInsert
    @ObjectName = 'helpActionsEdit'
   ,@Title = 'Amend Action'
   ,@TitleParameters = 'SystemID||ObjectName||ActionName'
   ,@HelpPage = 'helpActionsEdit.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpActionsEdit'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpActionsEdit'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsEdit'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsEdit'
   ,@ParameterName = 'ObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsEdit'
   ,@ParameterName = 'ActionName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpActionsEdit'
   ,@FieldName = 'HelpText'
   ,@Label = 'Description'
   ,@LabelWidth = 60
   ,@ValueType = 'string'
   ,@Width = 4000
   ,@DisplayWidth = 500
   ,@DisplayHeight = 20
   ,@Enabled = 'Y'
   ,@HelpText = 'Enter action description HTML.'
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsEdit'
   ,@ParameterName = 'AuditID'
   ,@ValueType = 'integer'
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'helpActionsEdit'
   ,@ActionName = 'Okay'
   ,@Process = 'helpActionsUpdate'
   ,@RowBased = 'Y'
   ,@Validate = 'Y'
   ,@CloseObject = 'Y'
   ,@ImageFile = 'okay.gif'
   ,@ToolTip = 'Save changes and exit'
   ,@KeyCode = 13
go

execute dbo.shlActionsInsert
    @ObjectName = 'helpActionsEdit'
   ,@ActionName = 'Cancel'
   ,@CloseObject = 'Y'
   ,@ImageFile = 'cancel.gif'
   ,@ToolTip = 'Exit without saving changes'
   ,@KeyCode = 27
go

print '.oOo.'
go

print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlDialogFormInsert
    @ObjectName = 'helpColoursEdit'
   ,@Title = 'Amend Colour'
   ,@TitleParameters = 'SystemID||ObjectName'
   ,@HelpPage = 'helpColoursEdit.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpColoursEdit'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpColoursEdit'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpColoursEdit'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpColoursEdit'
   ,@ParameterName = 'ObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpColoursEdit'
   ,@FieldName = 'ColourValue'
   ,@Label = 'Value'
   ,@ValueType = 'string'
   ,@Width = 200
   ,@DisplayWidth = 200
   ,@DisplayType = 'L'
   ,@IsPrimary = 'Y'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpColoursEdit'
   ,@FieldName = 'ValueDescription'
   ,@Label = 'Description'
   ,@ValueType = 'string'
   ,@Width = 30
   ,@DisplayWidth = 150
   ,@Enabled = 'Y'
   ,@HelpText = 'Enter description for this colour value.'
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpColoursEdit'
   ,@ParameterName = 'AuditID'
   ,@ValueType = 'integer'
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'helpColoursEdit'
   ,@ActionName = 'Okay'
   ,@Process = 'helpColoursUpdate'
   ,@RowBased = 'Y'
   ,@Validate = 'Y'
   ,@CloseObject = 'Y'
   ,@ImageFile = 'okay.gif'
   ,@ToolTip = 'Save changes and exit'
   ,@KeyCode = 13
go

execute dbo.shlActionsInsert
    @ObjectName = 'helpColoursEdit'
   ,@ActionName = 'Cancel'
   ,@CloseObject = 'Y'
   ,@ImageFile = 'cancel.gif'
   ,@ToolTip = 'Exit without saving changes'
   ,@KeyCode = 27
go

print '.oOo.'
go

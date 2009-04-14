print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlDialogFormInsert
    @ObjectName = 'helpObjectEdit'
   ,@Title = 'Amend Object'
   ,@TitleParameters = 'SystemID||ObjectName'
   ,@HelpPage = 'helpObjectEdit.html'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpObjectEdit'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpObjectEdit'
go

---------------------------------------------------

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObjectEdit'
   ,@FieldName = 'SystemID'
   ,@Label = 'System'
   ,@ValueType = 'string'
   ,@Width = 12
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
   ,@IsPrimary = 'Y'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObjectEdit'
   ,@FieldName = 'ObjectName'
   ,@Label = 'Object'
   ,@ValueType = 'string'
   ,@Width = 32
   ,@DisplayWidth = -1
   ,@DisplayType = 'H'
   ,@IsPrimary = 'Y'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObjectEdit'
   ,@FieldName = 'HelpText'
   ,@Label = 'Description'
   ,@LabelWidth = 60
   ,@ValueType = 'string'
   ,@Width = 4000
   ,@DisplayWidth = 500
   ,@DisplayHeight = 15
   ,@Enabled = 'Y'
   ,@HelpText = 'Enter the object description HTML.'
go

execute dbo.shlFieldParamInsert
    @ObjectName = 'helpObjectEdit'
   ,@FieldName = 'ColourText'
   ,@Label = 'Colour'
   ,@LabelWidth = 60
   ,@ValueType = 'string'
   ,@Width = 2000
   ,@DisplayWidth = 500
   ,@DisplayHeight = 5
   ,@Enabled = 'Y'
   ,@HelpText = 'Enter the object colour table description HTML.'
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpObjectEdit'
   ,@ParameterName = 'AuditID'
   ,@ValueType = 'integer'
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'helpObjectEdit'
   ,@ActionName = 'Okay'
   ,@Process = 'helpObjectUpdate'
   ,@RowBased = 'Y'
   ,@Validate = 'Y'
   ,@CloseObject = 'Y'
   ,@ImageFile = 'okay.gif'
   ,@ToolTip = 'Save changes and exit'
   ,@KeyCode = 13
go

execute dbo.shlActionsInsert
    @ObjectName = 'helpObjectEdit'
   ,@ActionName = 'Cancel'
   ,@CloseObject = 'Y'
   ,@ImageFile = 'cancel.gif'
   ,@ToolTip = 'Exit without saving changes'
   ,@KeyCode = 27
go

print '.oOo.'
go

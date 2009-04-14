print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @objectname = 'helpActionsGet'
   ,@procname = 'helpActionsGet'
   ,@dataparameter = 'helpActionsGet'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpActionsGet'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpActionsGet'
   ,@UpdateParent = 'Y'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsGet'
   ,@ParameterName = 'pSystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsGet'
   ,@ParameterName = 'pObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsGet'
   ,@ParameterName = 'pActionName'
   ,@ValueType = 'string'
   ,@Width = 32
go

print '.oOo.'
go

print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @objectname = 'helpActionsAuditGet'
   ,@procname = 'helpActionsAuditGet'
   ,@dataparameter = 'helpActionsAuditGet'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpActionsAuditGet'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpActionsAuditGet'
   ,@UpdateParent = 'Y'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsAuditGet'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsAuditGet'
   ,@ParameterName = 'ObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsAuditGet'
   ,@ParameterName = 'ActionName'
   ,@ValueType = 'string'
   ,@Width = 32
go

print '.oOo.'
go

print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @objectname = 'helpObjectAuditGet'
   ,@procname = 'helpObjectsAuditGet'
   ,@dataparameter = 'helpObjectAuditGet'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpObjectAuditGet'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpObjectAuditGet'
   ,@UpdateParent = 'Y'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpObjectAuditGet'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpObjectAuditGet'
   ,@ParameterName = 'ObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

print '.oOo.'
go

print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @objectname = 'SystemsAuditGet'
   ,@procname = 'helpSystemsAuditGet'
   ,@dataparameter = 'SystemsAuditGet'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'SystemsAuditGet'
   ,@ModuleID = 'helpSystem'
   ,@ObjectName = 'SystemsAuditGet'
   ,@UpdateParent = 'Y'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsAuditGet'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

print '.oOo.'
go

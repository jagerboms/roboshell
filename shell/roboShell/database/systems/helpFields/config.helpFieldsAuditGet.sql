print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @objectname = 'helpFieldsAuditGet'
   ,@procname = 'helpFieldsAuditGet'
   ,@dataparameter = 'helpFieldsAuditGet'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpFieldsAuditGet'
   ,@ModuleID = 'helpSystem'
   ,@ObjectName = 'helpFieldsAuditGet'
   ,@UpdateParent = 'Y'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsAuditGet'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsAuditGet'
   ,@ParameterName = 'ObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsAuditGet'
   ,@ParameterName = 'FieldName'
   ,@ValueType = 'string'
   ,@Width = 32
go

print '.oOo.'
go

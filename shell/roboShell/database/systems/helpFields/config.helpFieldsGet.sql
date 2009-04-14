print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @objectname = 'helpFieldsGet'
   ,@procname = 'helpFieldsGet'
   ,@dataparameter = 'helpFieldsGet'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpFieldsGet'
   ,@ModuleID = 'helpSystem'
   ,@ObjectName = 'helpFieldsGet'
   ,@UpdateParent = 'Y'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsGet'
   ,@ParameterName = 'pSystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsGet'
   ,@ParameterName = 'pObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsGet'
   ,@ParameterName = 'pFieldName'
   ,@ValueType = 'string'
   ,@Width = 32
go

print '.oOo.'
go

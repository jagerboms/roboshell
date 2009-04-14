print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @objectname = 'SystemsGet'
   ,@procname = 'helpSystemsGet'
   ,@dataparameter = 'SystemsGet'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'SystemsGet'
   ,@ModuleID = 'helpSystem'
   ,@ObjectName = 'SystemsGet'
   ,@UpdateParent = 'Y'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsGet'
   ,@ParameterName = 'pSystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

print '.oOo.'
go

print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @objectname = 'SystemsList'
   ,@procname = 'helpSystemsList'
   ,@dataparameter = 'SystemsList'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'SystemsList'
   ,@ModuleID = 'public'
   ,@ObjectName = 'SystemsList'
   ,@UpdateParent = 'Y'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsList'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

print '.oOo.'
go

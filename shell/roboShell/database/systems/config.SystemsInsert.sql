print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @objectname = 'SystemsInsert'
   ,@procname = 'helpSystemsInsert'
   ,@mode = 'X'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'SystemsInsert'
   ,@ModuleID = 'helpSystemMaintain'
   ,@ObjectName = 'SystemsInsert'
   ,@SuccessProcess = 'SystemsGet'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsInsert'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsInsert'
   ,@ParameterName = 'SystemName'
   ,@ValueType = 'string'
   ,@Width = 100
go

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsInsert'
   ,@ParameterName = 'Copyright'
   ,@ValueType = 'string'
   ,@Width = 100
go

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsInsert'
   ,@ParameterName = 'SystemsGet'
   ,@ValueType = 'object'
   ,@IsInput = 'N'
go

print '.oOo.'
go

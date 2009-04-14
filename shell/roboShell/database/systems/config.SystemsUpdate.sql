print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @objectname = 'SystemsUpdate'
   ,@procname = 'helpSystemsUpdate'
   ,@mode = 'P'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'SystemsUpdate'
   ,@ModuleID = 'helpSystemMaintain'
   ,@ObjectName = 'SystemsUpdate'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsUpdate'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsUpdate'
   ,@ParameterName = 'SystemName'
   ,@ValueType = 'string'
   ,@Width = 100
go

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsUpdate'
   ,@ParameterName = 'Copyright'
   ,@ValueType = 'string'
   ,@Width = 100
go

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsUpdate'
   ,@ParameterName = 'AuditID'
   ,@ValueType = 'integer'
go

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsUpdate'
   ,@ParameterName = 'StateName'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@IsInput = 'N'
go

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsUpdate'
   ,@ParameterName = 'State'
   ,@ValueType = 'string'
   ,@Width = 2
   ,@IsInput = 'N'
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'SystemsUpdate'
   ,@PropertyType = 'sk'
   ,@PropertyName = 'SYSTEM'
   ,@Value = ''
go

print '.oOo.'
go

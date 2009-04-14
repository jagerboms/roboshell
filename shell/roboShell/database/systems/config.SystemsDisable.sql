print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @objectname = 'SystemsDisable'
   ,@procname = 'helpSystemsDisable'
   ,@mode = 'P'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'SystemsDisable'
   ,@ModuleID = 'helpSystem'
   ,@ObjectName = 'SystemsDisable'
   ,@ConfirmMsg = 'Do you wish to disable this System?'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsDisable'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsDisable'
   ,@ParameterName = 'AuditID'
   ,@ValueType = 'integer'
go

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsDisable'
   ,@ParameterName = 'StateName'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@IsInput = 'N'
go

execute dbo.shlParametersInsert
    @ObjectName = 'SystemsDisable'
   ,@ParameterName = 'State'
   ,@ValueType = 'string'
   ,@Width = 2
   ,@IsInput = 'N'
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'SystemsDisable'
   ,@PropertyType = 'sk'
   ,@PropertyName = 'SYSTEM'
   ,@Value = ''
go

print '.oOo.'
go

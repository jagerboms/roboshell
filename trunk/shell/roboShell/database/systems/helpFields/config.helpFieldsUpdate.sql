print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @objectname = 'helpFieldsUpdate'
   ,@procname = 'helpFieldsUpdate'
   ,@mode = 'P'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpFieldsUpdate'
   ,@ModuleID = 'helpSystem'
   ,@ObjectName = 'helpFieldsUpdate'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsUpdate'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsUpdate'
   ,@ParameterName = 'ObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsUpdate'
   ,@ParameterName = 'FieldName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsUpdate'
   ,@ParameterName = 'HelpText'
   ,@ValueType = 'string'
   ,@Width = 4000
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsUpdate'
   ,@ParameterName = 'AuditID'
   ,@ValueType = 'integer'
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsUpdate'
   ,@ParameterName = 'StateName'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@IsInput = 'N'
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpFieldsUpdate'
   ,@ParameterName = 'State'
   ,@ValueType = 'string'
   ,@Width = 2
   ,@IsInput = 'N'
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpFieldsUpdate'
   ,@PropertyType = 'sk'
   ,@PropertyName = 'HELPFIELD'
   ,@Value = ''
go

print '.oOo.'
go

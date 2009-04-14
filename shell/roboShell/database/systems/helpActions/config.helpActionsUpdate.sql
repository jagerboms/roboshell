print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @objectname = 'helpActionsUpdate'
   ,@procname = 'helpActionsUpdate'
   ,@mode = 'P'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpActionsUpdate'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpActionsUpdate'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsUpdate'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsUpdate'
   ,@ParameterName = 'ObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsUpdate'
   ,@ParameterName = 'ActionName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsUpdate'
   ,@ParameterName = 'HelpText'
   ,@ValueType = 'string'
   ,@Width = 4000
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsUpdate'
   ,@ParameterName = 'AuditID'
   ,@ValueType = 'integer'
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsUpdate'
   ,@ParameterName = 'StateName'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@IsInput = 'N'
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpActionsUpdate'
   ,@ParameterName = 'State'
   ,@ValueType = 'string'
   ,@Width = 2
   ,@IsInput = 'N'
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpActionsUpdate'
   ,@PropertyType = 'sk'
   ,@PropertyName = 'HELPACTION'
   ,@Value = ''
go

print '.oOo.'
go

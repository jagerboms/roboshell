print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @objectname = 'helpColoursUpdate'
   ,@procname = 'helpColoursUpdate'
   ,@mode = 'P'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpColoursUpdate'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpColoursUpdate'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpColoursUpdate'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpColoursUpdate'
   ,@ParameterName = 'ObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpColoursUpdate'
   ,@ParameterName = 'ColourValue'
   ,@ValueType = 'string'
   ,@Width = 200
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpColoursUpdate'
   ,@ParameterName = 'ValueDescription'
   ,@ValueType = 'string'
   ,@Width = 30
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpColoursUpdate'
   ,@ParameterName = 'AuditID'
   ,@ValueType = 'integer'
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpColoursUpdate'
   ,@ParameterName = 'StateName'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@IsInput = 'N'
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpColoursUpdate'
   ,@ParameterName = 'State'
   ,@ValueType = 'string'
   ,@Width = 2
   ,@IsInput = 'N'
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpColoursUpdate'
   ,@PropertyType = 'sk'
   ,@PropertyName = 'HELPCOLOUR'
   ,@Value = ''
go

print '.oOo.'
go

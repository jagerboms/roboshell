print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @objectname = 'helpObjectUpdate'
   ,@procname = 'helpObjectsUpdate'
   ,@mode = 'P'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpObjectUpdate'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpObjectUpdate'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpObjectUpdate'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpObjectUpdate'
   ,@ParameterName = 'ObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpObjectUpdate'
   ,@ParameterName = 'HelpText'
   ,@ValueType = 'string'
   ,@Width = 4000
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpObjectUpdate'
   ,@ParameterName = 'ColourText'
   ,@ValueType = 'string'
   ,@Width = 2000
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpObjectUpdate'
   ,@ParameterName = 'AuditID'
   ,@ValueType = 'integer'
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpObjectUpdate'
   ,@ParameterName = 'StateName'
   ,@ValueType = 'string'
   ,@Width = 50
   ,@IsInput = 'N'
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpObjectUpdate'
   ,@ParameterName = 'State'
   ,@ValueType = 'string'
   ,@Width = 2
   ,@IsInput = 'N'
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'helpObjectUpdate'
   ,@PropertyType = 'sk'
   ,@PropertyName = 'HELPOBJECT'
   ,@Value = ''
go

print '.oOo.'
go

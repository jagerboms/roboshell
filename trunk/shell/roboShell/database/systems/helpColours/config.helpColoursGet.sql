print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @objectname = 'helpColoursGet'
   ,@procname = 'helpColoursGet'
   ,@dataparameter = 'helpColoursGet'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpColoursGet'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpColoursGet'
   ,@UpdateParent = 'Y'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpColoursGet'
   ,@ParameterName = 'pSystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpColoursGet'
   ,@ParameterName = 'pObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpColoursGet'
   ,@ParameterName = 'pColourValue'
   ,@ValueType = 'string'
   ,@Width = 200
go

print '.oOo.'
go

print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @objectname = 'helpColoursAuditGet'
   ,@procname = 'helpColoursAuditGet'
   ,@dataparameter = 'helpColoursAuditGet'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpColoursAuditGet'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpColoursAuditGet'
   ,@UpdateParent = 'Y'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpColoursAuditGet'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpColoursAuditGet'
   ,@ParameterName = 'ObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpColoursAuditGet'
   ,@ParameterName = 'ColourValue'
   ,@ValueType = 'string'
   ,@Width = 200
go

print '.oOo.'
go

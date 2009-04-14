print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @objectname = 'helpObjectGet'
   ,@procname = 'helpObjectsGet'
   ,@dataparameter = 'helpObjectGet'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'helpObjectGet'
   ,@ModuleID = 'helpsystem'
   ,@ObjectName = 'helpObjectGet'
   ,@UpdateParent = 'Y'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'helpObjectGet'
   ,@ParameterName = 'pSystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'helpObjectGet'
   ,@ParameterName = 'pObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

print '.oOo.'
go

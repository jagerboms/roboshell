print '-----------------------'
print '-- Akuna Care - Pets --'
print '-----------------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @objectname = 'shlUserPropertyAlter'
   ,@procname = 'shlUserPropertyAlter'
   ,@mode = 'X'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'shlUserPropertyAlter'
   ,@ModuleID = 'public'
   ,@ObjectName = 'shlUserPropertyAlter'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @objectname = 'shlUserPropertyAlter'
   ,@ParameterName = 'AddressType'
   ,@ValueType = 'string'
   ,@Width = 3
go

execute dbo.shlParametersInsert
    @objectname = 'shlUserPropertyAlter'
   ,@ParameterName = 'ObjectName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @objectname = 'shlUserPropertyAlter'
   ,@ParameterName = 'PropertyType'
   ,@ValueType = 'string'
   ,@Width = 2
go

execute dbo.shlParametersInsert
    @objectname = 'shlUserPropertyAlter'
   ,@ParameterName = 'PropertyName'
   ,@ValueType = 'string'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @objectname = 'shlUserPropertyAlter'
   ,@ParameterName = 'Value'
   ,@ValueType = 'string'
   ,@Width = 2000
go

print '.oOo.'
go

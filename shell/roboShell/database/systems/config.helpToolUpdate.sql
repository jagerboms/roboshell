print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlCallAsmInsert
    @ObjectName = 'helpToolUpdate'
   ,@LibraryName = 'HelpTool'
   ,@ClassName = 'HelpTool'
   ,@MethodName = 'UpdateHelpTables'
go

---------------------------------------------------

execute shlProcessesInsert
    @ProcessName = 'helpToolUpdate'
   ,@ModuleID = 'public'
   ,@ObjectName = 'helpToolUpdate'
go

---------------------------------------------------

execute shlParametersInsert
    @ObjectName = 'helpToolUpdate'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'String'
   ,@Width = 12
go

---------------------------------------------------

execute shlPropertiesInsert
    @ObjectName = 'helpToolUpdate'
   ,@PropertyType = 'mh'
   ,@PropertyName = 'SystemID'
   ,@Value = ''
go

print '.oOo.'
go

print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlCallAsmInsert
    @ObjectName = 'helpToolBuild'
   ,@LibraryName = 'HelpTool'
   ,@ClassName = 'HelpTool'
   ,@MethodName = 'BuildHelpPages'
go

---------------------------------------------------

execute shlProcessesInsert
    @ProcessName = 'helpToolBuildx'
   ,@ModuleID = 'public'
   ,@ObjectName = 'helpToolBuild'
go

execute shlProcessesInsert
    @ProcessName = 'helpToolBuild'
   ,@ModuleID = 'public'
   ,@ObjectName = 'helpBuildPath'
   ,@SuccessProcess = 'helpToolBuildx'
go

---------------------------------------------------

execute shlParametersInsert
    @ObjectName = 'helpToolBuild'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'String'
   ,@Width = 12
go

execute shlParametersInsert
    @ObjectName = 'helpToolBuild'
   ,@ParameterName = 'Path'
   ,@ValueType = 'String'
   ,@Width = 128
go

---------------------------------------------------

execute shlPropertiesInsert
    @ObjectName = 'helpToolBuild'
   ,@PropertyType = 'mh'
   ,@PropertyName = 'SystemID'
   ,@Value = ''
go

execute shlPropertiesInsert
    @ObjectName = 'helpToolBuild'
   ,@PropertyType = 'mh'
   ,@PropertyName = 'Path'
   ,@Value = ''
go

print '.oOo.'
go

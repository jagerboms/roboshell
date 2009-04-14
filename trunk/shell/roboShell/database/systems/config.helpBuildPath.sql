print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go

execute dbo.shlDirectoryInsert
    @ObjectName = 'helpBuildPath'
   ,@Title = 'Help files output directory'
   ,@OutputParameter = 'Path'
go

---------------------------------------------------

execute shlParametersInsert
    @ObjectName = 'helpBuildPath'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'String'
   ,@Width = 12
go

execute shlParametersInsert
    @ObjectName = 'helpBuildPath'
   ,@ParameterName = 'Path'
   ,@ValueType = 'String'
   ,@Width = 128
go

print '.oOo.'
go

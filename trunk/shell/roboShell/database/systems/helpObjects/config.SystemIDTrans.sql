print '-------------------'
print '-- Bank Emulator --'
print '-------------------'
set nocount on
go

execute dbo.shlTransformInsert
    @objectname = 'SystemIDTrans'
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'SystemIDTrans'
   ,@ParameterName = 'SystemID'
   ,@ValueType = 'string'
   ,@Width = 12
go

execute dbo.shlParametersInsert
    @ObjectName = 'SystemIDTrans'
   ,@ParameterName = 'pSystemID'
   ,@ValueType = 'string'
   ,@Width = 12
   ,@IsInput = 'N'
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'SystemIDTrans'
   ,@PropertyType = 'tr'
   ,@PropertyName = 'SystemID'
   ,@Value = 'pSystemID'
go

print '.oOo.'
go

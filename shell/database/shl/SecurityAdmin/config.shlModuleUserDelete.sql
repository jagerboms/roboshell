print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @ObjectName = 'shlModuleUserDelete'
   ,@ProcName = 'shlModuleUserDelete'
   ,@Mode = 'X'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'shlModuleUserDelete'
   ,@ModuleID = 'securityadmin'
   ,@ObjectName = 'shlModuleUserDelete'
   ,@ConfirmMsg = 'Do you wish to remove this permission rule?'
   ,@dbo = 'Y'
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlModuleUserDelete'
   ,@PropertyType = 'sk'
   ,@PropertyName = 'SECADMINTREE'
   ,@Value = ''
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'shlModuleUserDelete'
   ,@ParameterName = 'UserName'
   ,@ValueType = 'String'
   ,@Width = 128
go

execute dbo.shlParametersInsert
    @ObjectName = 'shlModuleUserDelete'
   ,@ParameterName = 'ModuleID'
   ,@ValueType = 'String'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'shlModuleUserDelete'
   ,@ParameterName = 'Parent'
   ,@ValueType = 'String'
   ,@Width = 32
   ,@IsInput = 'N'
go

execute dbo.shlParametersInsert
    @ObjectName = 'shlModuleUserDelete'
   ,@ParameterName = 'Type'
   ,@ValueType = 'String'
   ,@Width = 2
   ,@IsInput = 'N'
go

print '.oOo.'
go

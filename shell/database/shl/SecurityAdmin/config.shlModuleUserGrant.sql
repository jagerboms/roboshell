print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @ObjectName = 'shlModuleUserGrant'
   ,@ProcName = 'shlModuleUserGrant'
   ,@Mode = 'X'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'shlModuleUserGrant'
   ,@ModuleID = 'securityadmin'
   ,@ObjectName = 'shlModuleUserGrant'
   ,@ConfirmMsg = 'Do you wish to grant this permission?'
   ,@dbo = 'Y'
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlModuleUserGrant'
   ,@PropertyType = 'sk'
   ,@PropertyName = 'SECADMINTREE'
   ,@Value = ''
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'shlModuleUserGrant'
   ,@ParameterName = 'UserName'
   ,@ValueType = 'String'
   ,@Width = 128
go

execute dbo.shlParametersInsert
    @ObjectName = 'shlModuleUserGrant'
   ,@ParameterName = 'ModuleID'
   ,@ValueType = 'String'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'shlModuleUserGrant'
   ,@ParameterName = 'Parent'
   ,@ValueType = 'String'
   ,@Width = 32
   ,@IsInput = 'N'
go

execute dbo.shlParametersInsert
    @ObjectName = 'shlModuleUserGrant'
   ,@ParameterName = 'Type'
   ,@ValueType = 'String'
   ,@Width = 2
   ,@IsInput = 'N'
go

print '.oOo.'
go

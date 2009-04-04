print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @ObjectName = 'shlModuleUserDeny'
   ,@ProcName = 'shlModuleUserDeny'
   ,@Mode = 'X'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'shlModuleUserDeny'
   ,@ModuleID = 'securityadmin'
   ,@ObjectName = 'shlModuleUserDeny'
   ,@ConfirmMsg = 'Do you wish to deny this permission?'
   ,@dbo = 'Y'
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlModuleUserDeny'
   ,@PropertyType = 'sk'
   ,@PropertyName = 'SECADMINTREE'
   ,@Value = ''
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'shlModuleUserDeny'
   ,@ParameterName = 'UserName'
   ,@ValueType = 'String'
   ,@Width = 128
go

execute dbo.shlParametersInsert
    @ObjectName = 'shlModuleUserDeny'
   ,@ParameterName = 'ModuleID'
   ,@ValueType = 'String'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'shlModuleUserDeny'
   ,@ParameterName = 'Parent'
   ,@ValueType = 'String'
   ,@Width = 32
   ,@IsInput = 'N'
go

execute dbo.shlParametersInsert
    @ObjectName = 'shlModuleUserDeny'
   ,@ParameterName = 'Type'
   ,@ValueType = 'String'
   ,@Width = 2
   ,@IsInput = 'N'
go

print '.oOo.'
go

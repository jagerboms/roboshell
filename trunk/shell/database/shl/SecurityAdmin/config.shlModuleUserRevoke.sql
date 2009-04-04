print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlStoredProcInsert
    @ObjectName = 'shlModuleUserRevoke'
   ,@ProcName = 'shlModuleUserRevoke'
   ,@Mode = 'X'
go

---------------------------------------------------

execute dbo.shlProcessesInsert
    @ProcessName = 'shlModuleUserRevoke'
   ,@ModuleID = 'securityadmin'
   ,@ObjectName = 'shlModuleUserRevoke'
   ,@ConfirmMsg = 'Do you wish to revoke this permission?'
   ,@dbo = 'Y'
go

---------------------------------------------------

execute dbo.shlPropertiesInsert
    @ObjectName = 'shlModuleUserRevoke'
   ,@PropertyType = 'sk'
   ,@PropertyName = 'SECADMINTREE'
   ,@Value = ''
go

---------------------------------------------------

execute dbo.shlParametersInsert
    @ObjectName = 'shlModuleUserRevoke'
   ,@ParameterName = 'UserName'
   ,@ValueType = 'String'
   ,@Width = 128
go

execute dbo.shlParametersInsert
    @ObjectName = 'shlModuleUserRevoke'
   ,@ParameterName = 'ModuleID'
   ,@ValueType = 'String'
   ,@Width = 32
go

execute dbo.shlParametersInsert
    @ObjectName = 'shlModuleUserRevoke'
   ,@ParameterName = 'Parent'
   ,@ValueType = 'String'
   ,@Width = 32
   ,@IsInput = 'N'
go

execute dbo.shlParametersInsert
    @ObjectName = 'shlModuleUserRevoke'
   ,@ParameterName = 'Type'
   ,@ValueType = 'String'
   ,@Width = 2
   ,@IsInput = 'N'
go

print '.oOo.'
go

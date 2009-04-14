print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlMenuInsert
    @ObjectName = 'MainMenu'      -- Main Menu
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'MainMenu'
   ,@ActionName = 'Statics'
   ,@ImageFile = 'config.gif'
   ,@MenuType = 'S'
   ,@ToolTip = 'Static maintenance'
go

execute dbo.shlActionsInsert
    @ObjectName = 'MainMenu'
   ,@ActionName = 'Systems'
   ,@Process = 'Systems'
   ,@MenuType = 'I'
   ,@Parent = 'Statics'
   ,@MenuText = 'Systems'
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'MainMenu'
   ,@ActionName = 'Services'
   ,@ImageFile = 'service.gif'
   ,@MenuType = 'S'
   ,@ToolTip = 'Services'
go

execute dbo.shlActionsInsert
    @ObjectName = 'MainMenu'
   ,@ActionName = 'Users'
   ,@Process = 'shlSecurityAdmin'
   ,@MenuType = 'I'
   ,@Parent = 'Services'
   ,@ImageFile = 'security.gif'
   ,@MenuText = 'Users'
go

---------------------------------------------------

print '.oOo.'
go

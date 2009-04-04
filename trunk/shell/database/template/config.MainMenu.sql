print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

execute dbo.shlMenuInsert
    @ObjectName = 'MainMenu'      -- Main Menu
go

execute dbo.shlActionsInsert
    @ObjectName = 'MainMenu'
   ,@ActionName = 'Web'
   ,@ImageFile = 'web.gif'
   ,@MenuType = 'S'
   ,@ToolTip = 'Web Maintainence'
go

---------------------------------------------------

execute dbo.shlActionsInsert
    @ObjectName = 'MainMenu'
   ,@ActionName = 'WebPages'
   ,@Process = 'webPages'
   ,@MenuType = 'I'
   ,@Parent = 'Web'
   ,@MenuText = 'Pages'
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

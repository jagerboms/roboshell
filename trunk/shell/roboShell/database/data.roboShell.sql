print '-------------------------'
print '-- Tolbeam Pty Limited --'
print '-------------------------'
set nocount on
go

if dbo.shlVariableGet('SystemName') is null
begin
    execute dbo.shlVariableSet
        @VariableID='SystemName', @VariableValue='Robo Shell', @ShellUse='Y'
    print dbo.shlVariableGet('SystemName')
end
go
if dbo.shlVariableGet('Release') is null
begin
    execute dbo.shlVariableSet
        @VariableID='Release', @VariableValue='0.1', @ShellUse='Y'
    print dbo.shlVariableGet('Release')
end
go
if dbo.shlVariableGet('Environment') is null
begin
    execute dbo.shlVariableSet
        @VariableID='Environment', @VariableValue='Dev', @ShellUse='Y'
    print dbo.shlVariableGet('Environment')
end
go
if dbo.shlVariableGet('Production') is null
begin
    execute dbo.shlVariableSet
        @VariableID='Production', @VariableValue='Y', @ShellUse='Y'
    print dbo.shlVariableGet('Production')
end
go

print '.oOo.'
go

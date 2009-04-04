print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlVariableGet') is not null
begin
    drop function dbo.shlVariableGet
end
go

create function dbo.shlVariableGet
(
    @VariableID varchar(32)
)
returns varchar(200)
as
begin
    declare @VariableValue varchar(200)

    select  @VariableValue = k.VariableValue
    from    dbo.shlVariables k
    where   k.VariableID = @VariableID
    return (@VariableValue)
end
go

print '.oOo.'
go

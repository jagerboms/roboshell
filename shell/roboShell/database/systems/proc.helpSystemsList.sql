print '----------------'
print '-- Robo Shell --'
print '----------------'
set nocount on
go
if object_id('dbo.helpSystemsList') is not null
begin
    drop procedure dbo.helpSystemsList
end
go
create procedure dbo.helpSystemsList
    @SystemID varchar(12) = null
as
begin
    set nocount on

    select  a.SystemID
           ,a.SystemName
    from    dbo.helpSystems a
    where   a.State = 'ac' or a.SystemID = @SystemID
    order by a.SystemName
end
go

print '.oOo.'
go

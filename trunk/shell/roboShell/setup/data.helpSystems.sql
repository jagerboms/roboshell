if not exists (select 'a' from dbo.helpSystems where SystemID = 'default')
begin
    execute dbo.helpSystemsInsert
        @SystemID = 'default'
       ,@SystemName = 'default help text definitions'
       ,@Copyright = 'Russell Hansen, Tolbeam Pty Limited'
end
go

if not exists (select 'a' from dbo.helpSystems where SystemID = 'roboShell')
begin
    execute dbo.helpSystemsInsert
        @SystemID = 'roboShell'
       ,@SystemName = 'roboShell tool'
       ,@Copyright = 'Russell Hansen, Tolbeam Pty Limited'
end
go

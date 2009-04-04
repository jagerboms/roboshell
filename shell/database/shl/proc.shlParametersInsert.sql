print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlParametersInsert') is not null
begin
    drop procedure dbo.shlParametersInsert
end
go

create procedure dbo.shlParametersInsert
    @ObjectName varchar(32)
   ,@ParameterName varchar(32)
   ,@ValueType varchar(25)
   ,@Width integer = 0
   ,@Value varchar(200) = null
   ,@IsInput char(1) = 'Y'
   ,@IsOutput char(1) = 'Y'
   ,@Type char(1) = 'U'
as
begin
    set nocount on
    declare @e integer
           ,@count integer
           ,@sequence integer

    set @e = 0
    while @e = 0
    begin
        print 'Parameter: ' + rtrim(@ObjectName) + '.' + @ParameterName

        set @IsInput = upper(@IsInput)
        if @IsInput <> 'Y'
        begin
            set @IsInput = 'N'
        end

        set @IsOutput = upper(@IsOutput)
        if @IsOutput <> 'Y'
        begin
            set @IsOutput = 'N'
        end

        select  @sequence = max(Sequence)
        from    dbo.shlParameters p
        where   ObjectName = @ObjectName

        set @sequence = coalesce(@sequence, 0) + 1

        insert into dbo.shlParameters
        (
            ObjectName, ParameterName, Sequence,
            IsInput, IsOutput, ValueType, Width,
            Value, Type
        )
        values
        (
            @ObjectName, @ParameterName, @Sequence,
            @IsInput, @IsOutput, @ValueType, @Width,
            @Value, @Type
        )
        set @e = @@error
        break
    end
    return @e
end
go

print '.oOo.'
go

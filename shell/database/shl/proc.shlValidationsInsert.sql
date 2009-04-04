print '-----------------------------------'
print 'Copyright Tolbeam Pty Limited'
print 'Application shell'
print '-----------------------------------'
set nocount on
go

if object_id('dbo.shlValidationsInsert') is not null
begin
    drop procedure dbo.shlValidationsInsert
end
go

create procedure dbo.shlValidationsInsert
    @ObjectName varchar(32)
   ,@ValidationName varchar(32)
   ,@FieldName varchar(32)
   ,@ValidationType char(2) = 'EQ' -- EQ,NE,NN,GT,GE,LT,LE
   ,@ValueType char(1) = 'C'       -- Constant, Field, Process
   ,@Process varchar(32) = null
   ,@Value varchar(200)
   ,@Message varchar(200)
   ,@ReturnParameter varchar(32) = null
as
begin
    set nocount on
    declare @e integer
           ,@m varchar(200)

    set @e = 0
    while @e = 0
    begin
        print 'Validation: ' + rtrim(@ObjectName) + '.' + rtrim(@ValidationName)

        if @ReturnParameter is not null
        begin
            if not exists
            (
                select  'a'
                from    dbo.shlParameters p
                where   p.ObjectName = @ObjectName
                and     p.ParameterName = @ReturnParameter
            )
            begin
                set @e = 50041
                set @m = 'Error: parameter ' + Coalesce(@ReturnParameter, 'null') 
                                + ' does not exist in the database'
                raiserror @e @m
                break
            end
        end
        else
        begin
            if @ValueType = 'P'
            begin
                set @e = 50042
                set @m = 'Error: return parameter must be provided for Process value type'
                raiserror @e @m
                break
            end
        end

        if @Process is not null
        begin
            if not exists
            (
                select  'a'
                from    dbo.shlProcesses p
                where   p.ProcessName = @Process
            )
            begin
                set @e = 50043
                set @m = 'Error: process ' + Coalesce(@Process, 'null') 
                                + ' does not exist in the database'
                raiserror @e @m
                break
            end
        end
        else
        begin
            if @ValueType = 'P'
            begin
                set @e = 50044
                set @m = 'Error: Process must be provided for Process value type'
                raiserror @e @m
                break
            end
        end

        insert into dbo.shlValidations
        (
            ObjectName, ValidationName, FieldName,
            ValidationType, ValueType, Process,
            Value, Message, ReturnParameter
        )
        values
        (
            @ObjectName, @ValidationName, @FieldName,
            @ValidationType, @ValueType, @Process,
            @Value, @Message, @ReturnParameter
        )
        set @e = @@error
        break
    end
    return @e
end
go

print '.oOo.'
go

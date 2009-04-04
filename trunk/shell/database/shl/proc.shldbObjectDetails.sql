if object_id('dbo.shldbObjectDetails') is not null
begin
    drop procedure dbo.shldbObjectDetails
end
go
create procedure dbo.shldbObjectDetails
    @Name sysname
as
begin
    set nocount on 
    declare
        @obj integer
       ,@n sysname
       ,@type char(2)
       ,@CName sysname

    if coalesce(parsename(@Name,3), db_name()) <> db_name()
    begin
        raiserror(15250, -1, -1)
        return (1)
    end

    set @n = parsename(@Name, 1)
    select  @obj = s.id
           ,@type = s.type
    from    dbo.sysobjects s
    where   s.Name = @n

    if @@rowcount = 0
    begin
        set @CName = db_name()
        raiserror (15009, -1, -1, @n, @CName)
        return (1)
    end

    select  @n ObjectName
           ,@type Type

    if @type = 'U'
    begin
        select  COLUMN_NAME Name
               ,DATA_TYPE type
               ,CHARACTER_MAXIMUM_LENGTH length
               ,substring(IS_NULLABLE,1,1) nul
               ,NUMERIC_PRECISION xprec
               ,NUMERIC_SCALE scale
        from    INFORMATION_SCHEMA.COLUMNS
        where   TABLE_NAME = @n
        order by ORDINAL_POSITION

        select  x.name IndexName
               ,i.keyno KeyOrder
               ,index_col(object_name(x.id), x.indid, i.keyno) ColumnName
               ,case indexkey_property(x.id, x.indid, i.colid, 'isdescending') when 1 then 'Y' else 'N' end Descending
               ,case when indexproperty(x.id, x.name, 'IsClustered') = 1 then 'Y' else 'N' end Cluster
               ,case when s.name is not null then 'Y' else 'N' end PrimaryKey
               ,case when indexproperty(x.id, x.name, 'IsUnique') = 1 then 'Y' else 'N' end UniqueIndex
        from    dbo.sysindexes x
        join    dbo.sysindexkeys i
        on      i.id = x.id
        and     i.indid = x.indid
        left join dbo.sysobjects s 
    	on      s.name = x.name
    	and     s.parent_obj = x.id
        and     s.xtype = 'PK' 
        where   x.id = @obj
        and     indexproperty(x.id, x.name, 'IsStatistics') = 0
        order by 1, 2, 3

        select  col_name(@obj, s.info) ColumnName
               ,s.name ConstraintName
               ,c.text "default"
        from    dbo.sysobjects s 
        join    dbo.syscomments c
        on      c.id = s.id
        and     c.colid = 1
        where   s.parent_obj = @obj 
        and     s.xtype = 'D ' 

        select  object_name(k.constid) ConstraintName
               ,k.keyno Sequence
               ,col_name(k.fkeyid, k.fkey) ColumnName
               ,object_name(k.rkeyid) LinkedTable
               ,col_name(k.rkeyid, k.rkey) LinkedColumn
        from    dbo.sysforeignkeys k
        where   k.fkeyid = object_id(@n)
        order by 1, 2

        select  o.name TriggerName
        from    dbo.sysobjects o
        where   o.type = 'TR'
        and     o.parent_obj = @obj
    end
    else
    begin
        select  ORDINAL_POSITION ParameterOrder
               ,PARAMETER_NAME ParameterName
               ,DATA_TYPE type
               ,CHARACTER_MAXIMUM_LENGTH length
               ,NUMERIC_PRECISION xprec
               ,NUMERIC_SCALE scale
               ,PARAMETER_MODE InOut
        from    INFORMATION_SCHEMA.PARAMETERS
        where   SPECIFIC_NAME = @n
        order by ORDINAL_POSITION

        execute dbo.sp_helptext
            @objname = @n
    end
end
go 

print '.oOo.'
go

Option Explicit On
Option Strict On

#Region "copyright Russell Hansen, Tolbeam Pty Limited"
'dbSchema is free software issued as open source;
' you can redistribute it and/or modify it under the terms of the
' GNU General Public License version 2 as published by the Free Software Foundation.
'dbSchema is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
' without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
'See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with dbSchema;
' if not, go to the web site (http://www.gnu.org/licenses/gpl-2.0.html)
' or write to:
'   The Free Software Foundation, Inc.,
'   59 Temple Place,
'   Suite 330,
'   Boston, MA 02111-1307 USA. 
#End Region

Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Collections.Specialized

Public Class sql
    Private psConn As SqlConnection
    Private Connected As Boolean = False
    Private sServer As String = "(local)"
    Private sDatabase As String = "master"
    Private sUserID As String = ""
    Private sPassword As String = ""
    Private sNetwork As String = ""
    Private sConnect As String = ""
    Private iTimeOut As Integer = 0

#Region "Properties"
    Public Property Server() As String
        Get
            Server = sServer
        End Get
        Set(ByVal value As String)
            If sServer <> value Then
                sServer = value
                CloseConnect()
            End If
        End Set
    End Property

    Public Property Database() As String
        Get
            Database = sDatabase
        End Get
        Set(ByVal value As String)
            If sDatabase <> value Then
                sDatabase = value
                CloseConnect()
            End If
        End Set
    End Property

    Public Property UserID() As String
        Get
            UserID = sUserID
        End Get
        Set(ByVal value As String)
            If sUserID <> value Then
                sUserID = value
                CloseConnect()
            End If
        End Set
    End Property

    Public Property Password() As String
        Get
            Password = sPassword
        End Get
        Set(ByVal value As String)
            If sPassword <> value Then
                sPassword = value
                CloseConnect()
            End If
        End Set
    End Property

    Public Property Network() As String
        Get
            Network = sNetwork
        End Get
        Set(ByVal value As String)
            If sNetwork <> value Then
                sNetwork = value
                CloseConnect()
            End If
        End Set
    End Property

    Public Property ConnectString() As String
        Get
            ConnectString = getConnectString()
        End Get
        Set(ByVal value As String)
            If value <> getConnectString() Then
                sConnect = value
                sServer = ""
                sDatabase = ""
                sUserID = ""
                sPassword = ""
                sNetwork = ""
                CloseConnect()
            End If
        End Set
    End Property

    Public Property TimeOut() As Integer
        Get
            TimeOut = iTimeOut
        End Get
        Set(ByVal value As Integer)
            iTimeOut = value
        End Set
    End Property

    Public Version As Integer = -1
#End Region

#Region "Methods"
    Public Function DatabaseList() As DataTable
        Dim sql As String

        openConnect()
        sql = "select name,cmptlevel from master.dbo.sysdatabases where name not in ('master','tempdb','model','distribution')"

        Return GetTable(sql)
    End Function

    Public Function CheckAccess() As String
        Dim sql As String
        Dim dr As DataRow

        openConnect()
        sql = "select case when is_member('db_owner')=1 then 'Y' " & _
              "when is_member('db_ddladmin')= 1 then 'Y' " & _
              "when is_srvrolemember('sysadmin')=1 then 'Y' else 'N' end"

        dr = GetRow(sql)

        If Not dr Is Nothing Then
            If dr.Item(0).ToString = "Y" Then
                Return "OK"
            End If
        End If
        Return "NOK"
    End Function

    Public Function DatabaseObject(ByVal Name As String, ByVal Schema As String, _
                                ByVal Type As String) As DataTable
        Dim sql As String
        Dim s As String = ""
        Dim bT As Boolean = False
        Dim sC As String = ""

        Type = UCase(Type)
        If InStr(Type, "P") > 0 Then
            s = "'P'"
            sC = ","
        End If
        If InStr(Type, "U") > 0 Then
            s &= sC & "'U'"
            sC = ","
        End If
        If InStr(Type, "F") > 0 Then
            s &= sC & "'FN', 'TF'"
            sC = ","
        End If
        If InStr(Type, "V") > 0 Then
            s &= sC & "'V'"
            sC = ","
        End If
        If InStr(Type, "T") > 0 Then
            s &= sC & "'TR'"
            sC = ","
            bT = True
        End If
        If sC = "" Then
            s = "'P','U','FN','TF','V','TR'"
            bT = True
        End If

        openConnect()
        If Version < 90 Then            'SQL 2000 compatible
            sql = "select type,name,object_name(parent_obj) parent,user_name(uid) sch "
            sql &= "from dbo.sysobjects "
            sql &= "where type in (" & s & ") "
            If Name <> "" Then
                sql &= "and (name like '" & Name & "'"
                If bT Then
                    sql &= " or object_name(parent_obj) like '" & Name & "'"
                End If
                sql &= ") "
            End If
            If Schema <> "" Then
                sql &= "and user_name(uid)='" & Schema & "' "
            End If
            sql &= "and uid = 1 "
            sql &= "and name not like 'dt_%' "
            sql &= "and name not in ('syssegments','sysconstraints','sysdiagrams',"
            sql &= "'sp_alterdiagram','sp_creatediagram',"
            sql &= "'sp_dropdiagram','sp_helpdiagramdefinition','sp_helpdiagrams',"
            sql &= "'sp_renamediagram','sp_upgraddiagrams','fn_diagramobjects') "
            sql &= "order by type, name"
        Else
            sql = "select type,name,object_name(parent_object_id) parent,schema_name(schema_id) sch "
            sql &= "from sys.objects "
            sql &= "where type in (" & s & ") "
            If Name <> "" Then
                sql &= "and (name like '" & Name & "'"
                If bT Then
                    sql &= " or object_name(parent_object_id) like '" & Name & "'"
                End If
                sql &= ") "
            End If
            If Schema <> "" Then
                sql &= "and schema_name(schema_id)='" & Schema & "' "
            End If
            sql &= "and name not like 'dt_%' "
            sql &= "and name not in ('syssegments','sysconstraints','sysdiagrams',"
            sql &= "'sp_alterdiagram','sp_creatediagram',"
            sql &= "'sp_dropdiagram','sp_helpdiagramdefinition','sp_helpdiagrams',"
            sql &= "'sp_renamediagram','sp_upgraddiagrams','fn_diagramobjects') "
            sql &= "order by type, name"
        End If

        Return GetTable(sql)
    End Function

    Public Function ObjectText(ByVal Name As String, ByVal Schema As String) As DataTable
        Dim sql As String
        Dim s As String

        openConnect()

        s = Schema & "." & Name
        If Version < 90 Then            'SQL 2000 compatible
            sql = "select text from syscomments "
            sql &= "where id = object_id('" & s & "') order by number, colid"
        Else
            sql = "select text from sys.syscomments "
            sql &= "where id = object_id('" & s & "') order by number, colid"
        End If

        Return GetTable(sql)
    End Function

    Public Function ObjectSettings(ByVal Name As String, ByVal Schema As String) As DataTable
        Dim sql As String
        Dim s As String

        openConnect()

        s = Schema & "." & Name
        If Version < 90 Then            'SQL 2000 compatible
            sql = "select objectproperty(object_id('" & s & "'), 'IsAnsiNullsOn') nulls,"
            sql &= "objectproperty(object_id('" & s & "'), 'IsQuotedIdentOn') quoted"
        Else
            sql = "select uses_ansi_nulls nulls,uses_quoted_identifier quoted from sys.sql_modules"
            sql &= " where object_id=object_id('" & s & "')"
        End If

        Return GetTable(sql)
    End Function

    Public Function TableColumns(ByVal TableName As String, ByVal Schema As String) As DataTable
        Dim sql As String

        openConnect()

        If Version < 90 Then            'SQL 2000 compatible
            sql = "declare @z bit,@o bit set @z=0 set @o=1 "
            sql &= "select object_name(c.id) TableName"
            sql &= ",c.name COLUMN_NAME"
            sql &= ",t.name DATA_TYPE"
            sql &= ",c.prec CHARACTER_MAXIMUM_LENGTH"
            sql &= ",case c.isnullable when 0 then 'NO' else 'YES' end IS_NULLABLE"
            sql &= ",c.xprec NUMERIC_PRECISION"
            sql &= ",c.xscale NUMERIC_SCALE"
            sql &= ",s.name DEFAULT_NAME"
            sql &= ",cm.text DEFAULT_TEXT"
            sql &= ",c.collation COLLATION_NAME"
            sql &= ",case when c.status & 128 = 0 then 'YES' else 'NO' end ANSIPadded"
            sql &= ",'NO' Replicate"
            sql &= ",case when c.colstat & 2 = 0 then 'NO' else 'YES' end ROWGUID"
            sql &= ",m.text Computed"
            sql &= ",'NO' Persisted"
            sql &= ",@z is_xml_document"
            sql &= ",null xmlschema"
            sql &= ",null xmlcollection"
            sql &= ",case when c.colstat & 1=1 then @o else @z end is_identity"
            sql &= ",columnproperty(c.id,c.name,'IsIdNotForRepl') IdentityReplicated"
            sql &= ",case when c.colstat & 1=1 then ident_seed('" & Schema & "." & TableName & "') else 0 end IdentitySeed"
            sql &= ",case when c.colstat & 1=1 then ident_incr('" & Schema & "." & TableName & "') else 0 end IdentityIncrement "
            sql &= "from dbo.syscolumns c "
            sql &= "left join dbo.sysobjects s "
            sql &= "on s.id = c.cdefault "
            sql &= "left join dbo.syscomments cm "
            sql &= "on cm.id = s.id "
            sql &= "and cm.colid = 1 "
            sql &= "join dbo.systypes t "
            sql &= "on t.xtype = c.xtype "
            sql &= "and t.xusertype = c.xusertype "
            sql &= "left join dbo.syscomments m "
            sql &= "on m.id = c.id "
            sql &= "and m.number = c.colid "
            sql &= "where c.id = object_id('" & Schema & "." & TableName & "') "
            sql &= "order by c.colorder"
        Else
            sql = "select i.name TableName"
            sql &= ",c.name COLUMN_NAME"
            sql &= ",t.name DATA_TYPE"
            sql &= ",c.max_length / (case when t.name in ('nchar','nvarchar') then 2 else 1 end) CHARACTER_MAXIMUM_LENGTH"
            sql &= ",case c.is_nullable when 0 then 'NO' else 'YES' end IS_NULLABLE"
            sql &= ",c.precision NUMERIC_PRECISION"
            sql &= ",c.scale NUMERIC_SCALE"
            sql &= ",s.name DEFAULT_NAME"
            sql &= ",cm.text DEFAULT_TEXT"
            sql &= ",c.collation_name COLLATION_NAME"
            sql &= ",case c.is_ansi_padded when 0 then 'NO' else 'YES' end ANSIPadded"
            sql &= ",case c.is_replicated when 0 then 'NO' else 'YES' end Replicate"
            sql &= ",case c.is_rowguidcol when 0 then 'NO' else 'YES' end ROWGUID"
            sql &= ",m.definition Computed"
            sql &= ",case coalesce(m.is_persisted, 0) when 0 then 'NO' else 'YES' end Persisted"
            sql &= ",c.is_xml_document"
            sql &= ",sc.name xmlschema"
            sql &= ",x.name xmlcollection"
            sql &= ",c.is_identity"
            sql &= ",columnproperty(i.object_id,c.name,'IsIdNotForRepl') IdentityReplicated"
            sql &= ",case c.is_identity when 1 then ident_seed(schema_name(i.schema_id)+'.'+i.name) else 0 end IdentitySeed"
            sql &= ",case c.is_identity when 1 then ident_incr(schema_name(i.schema_id)+'.'+i.name) else 0 end IdentityIncrement "
            sql &= "from sys.objects i "
            sql &= "join sys.columns c "
            sql &= "on c.object_id = i.object_id "
            sql &= "left join sys.objects s "
            sql &= "on s.object_id = c.default_object_id "
            sql &= "left join dbo.syscomments cm "
            sql &= "on cm.id = s.object_id "
            sql &= "and cm.colid = 1 "
            sql &= "join dbo.systypes t "
            sql &= "on t.xtype = c.system_type_id "
            sql &= "and t.xusertype = c.user_type_id "
            sql &= "left join sys.computed_columns m "
            sql &= "on m.object_id = c.object_id "
            sql &= "and m.column_id = c.column_id "
            sql &= "left join sys.xml_schema_collections x "
            sql &= "on x.xml_collection_id = c.xml_collection_id "
            sql &= "left join sys.schemas sc "
            sql &= "on sc.schema_id = x.schema_id "
            sql &= "where i.object_id = object_id('" & Schema & "." & TableName & "') "
            sql &= "order by c.column_id"
        End If

        Return GetTable(sql)
    End Function

    Public Function TableDetails(ByVal TableName As String, ByVal Schema As String) As DataRow
        Dim sql As String
        Dim s As String

        openConnect()

        s = Schema & "." & TableName
        If Version < 90 Then            'SQL 2000 compatible
            sql = "declare @f sysname,@p sysname "
            sql &= "select @f=groupname from sysfilegroups where status & 16 <> 0 "
            sql &= "select @f DataFileGroup"
            sql &= ",@f TextFileGroup"
            sql &= ",@f DefFileGroup"
            sql &= ",@p PartitionScheme"
            sql &= ",@p SchemeColumn"
            sql &= ",databasepropertyex(db_name(), 'Collation') DefCollation"
        Else
            sql = "declare @n sysname,@d sysname,@t sysname,@f sysname,@p sysname,@c sysname "
            sql &= "set @n='" & s & "' "
            sql &= "select @d=d.name "
            sql &= "from sys.partitions p "
            sql &= "join sys.indexes i "
            sql &= "on i.object_id=p.object_id "
            sql &= "and i.index_id=p.index_id "
            sql &= "left join sys.data_spaces d "
            sql &= "on d.data_space_id=i.data_space_id "
            sql &= "where p.object_id=object_id(@n) "
            sql &= "and p.index_id < 2 "
            sql &= "select @t=d.name "
            sql &= "from sys.tables t "
            sql &= "left join sys.data_spaces d "
            sql &= "on d.data_space_id=t.lob_data_space_id "
            sql &= "where t.object_id=object_id(@n) "
            sql &= "select @p=pf.name"
            sql &= ",@c=case when pf.name is null then null else '?unknown?' end "
            sql &= "from sys.partitions p "
            sql &= "join sys.indexes i "
            sql &= "on i.object_id=p.object_id "
            sql &= "and i.index_id=p.index_id "
            sql &= "left join sys.partition_schemes ps "
            sql &= "on ps.data_space_id=i.data_space_id "
            sql &= "left join sys.partition_functions pf "
            sql &= "on pf.function_id=ps.function_id "
            sql &= "where p.object_id=object_id(@n) "
            sql &= "and p.index_id < 2 "
            sql &= "select @c=c.name "
            sql &= "from sys.partitions p "
            sql &= "join sys.index_columns i "
            sql &= "on i.object_id=p.object_id "
            sql &= "and i.index_id=p.index_id "
            sql &= "and i.partition_ordinal>0 "
            sql &= "join sys.columns c "
            sql &= "on c.object_id=i.object_id "
            sql &= "and c.column_id=i.column_id "
            sql &= "where p.object_id=object_id(@n) "
            sql &= "and p.index_id < 2 "
            sql &= "select @f=d.name from sys.data_spaces d where is_default=1 "
            sql &= "select @d DataFileGroup"
            sql &= ",@t TextFileGroup"
            sql &= ",@f DefFileGroup"
            sql &= ",@p PartitionScheme"
            sql &= ",@c SchemeColumn"
            sql &= ",databasepropertyex(db_name(), 'Collation') DefCollation"
        End If
        Return GetRow(sql)
    End Function

    Public Function TableFKeys(ByVal TableName As String, ByVal Schema As String) As DataTable
        Dim sql As String

        openConnect()

        If Version < 90 Then            'SQL 2000 compatible
            sql = "select c.CONSTRAINT_NAME ConstraintName"
            sql &= ",k.keyno Sequence"
            sql &= ",col_name(k.fkeyid, k.fkey) ColumnName"
            sql &= ",user_name(o1.uid) LinkedSchema"
            sql &= ",object_name(k.rkeyid) LinkedTable"
            sql &= ",col_name(k.rkeyid, k.rkey) LinkedColumn"
            sql &= ",c.MATCH_OPTION"
            sql &= ",c.UPDATE_RULE"
            sql &= ",c.DELETE_RULE"
            sql &= ",objectproperty(object_id(k.constid),'CnstIsNotRepl') Replicated "
            sql &= "from dbo.sysforeignkeys k "
            sql &= "join INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS c "
            sql &= "on c.CONSTRAINT_NAME = object_name(k.constid) "
            sql &= "join sysobjects o "
            sql &= "on o.id=k.fkeyid "
            sql &= "join sysobjects o1 "
            sql &= "on o1.id=k.rkeyid "
            sql &= "where o.name='" & TableName & "' "
            sql &= "and user_name(o.uid)='" & Schema & "' "
            sql &= "order by 1, 2"
        Else
            sql = "select c.CONSTRAINT_NAME ConstraintName"
            sql &= ",u1.ORDINAL_POSITION Sequence"
            sql &= ",u1.COLUMN_NAME ColumnName"
            sql &= ",u2.TABLE_SCHEMA LinkedSchema"
            sql &= ",u2.TABLE_NAME LinkedTable"
            sql &= ",u2.COLUMN_NAME LinkedColumn"
            sql &= ",c.MATCH_OPTION"
            sql &= ",c.UPDATE_RULE"
            sql &= ",c.DELETE_RULE"
            sql &= ",objectproperty(object_id(c.CONSTRAINT_SCHEMA+'.'+c.CONSTRAINT_NAME),'CnstIsNotRepl') Replicated "
            sql &= "from INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS c "
            sql &= "join INFORMATION_SCHEMA.KEY_COLUMN_USAGE u1 "
            sql &= "on u1.CONSTRAINT_CATALOG=c.CONSTRAINT_CATALOG "
            sql &= "and u1.CONSTRAINT_SCHEMA=c.CONSTRAINT_SCHEMA "
            sql &= "and u1.CONSTRAINT_NAME=c.CONSTRAINT_NAME "
            sql &= "join INFORMATION_SCHEMA.KEY_COLUMN_USAGE u2 "
            sql &= "on u2.CONSTRAINT_CATALOG=c.UNIQUE_CONSTRAINT_CATALOG "
            sql &= "and u2.CONSTRAINT_SCHEMA=c.UNIQUE_CONSTRAINT_SCHEMA "
            sql &= "and u2.CONSTRAINT_NAME=c.UNIQUE_CONSTRAINT_NAME "
            sql &= "and u2.ORDINAL_POSITION=u1.ORDINAL_POSITION "
            sql &= "where u1.TABLE_NAME='" & TableName & "' "
            sql &= "and u1.TABLE_SCHEMA='" & Schema & "' "
            sql &= "order by 1, 2"
        End If

        Return GetTable(sql)
    End Function

    Public Function TableIndexes(ByVal TableName As String, ByVal Schema As String) As DataTable
        Dim sql As String

        openConnect()

        If Version < 90 Then            'SQL 2000 compatible
            sql = "declare @z bit, @o bit set @z=0 set @o=1 "
            sql &= "select Case x.name"
            sql &= ",i.keyno index_column_id"
            sql &= ",index_col(object_name(x.id), x.indid, i.keyno) ColumnName"
            sql &= ",i.keyno key_ordinal"
            sql &= ",0 partition_ordinal"
            sql &= ",case when indexproperty(x.id, x.name, 'IsClustered') = 1 then 1 else 2 end type"
            sql &= ",case when s.name is not null then @o else @z end is_primary_key"
            sql &= ",case when indexproperty(x.id, x.name, 'IsUnique') = 1 then @o else @z end is_unique"
            sql &= ",case indexkey_property(x.id, x.indid, i.colid, 'isdescending') when 1 then @o else @z end is_descending_key"
            sql &= ",@z is_included_column"
            sql &= ",'PRIMARY' filegroup"
            sql &= ",x.OrigFillFactor FILL_FACTOR"
            sql &= ",@z PAD_INDEX"
            sql &= ",@z IGNORE_DUP_KEY"
            sql &= ",case x.lockflags & 1 when 0 then @o else @z end ALLOW_ROW_LOCKS "
            sql &= ",case x.lockflags & 2 when 0 then @o else @z end ALLOW_PAGE_LOCKS "
            sql &= ",@z no_recompute "
            sql &= "from dbo.sysindexes x "
            sql &= "join dbo.sysindexkeys i "
            sql &= "on i.id = x.id "
            sql &= "and i.indid = x.indid "
            sql &= "left join dbo.sysobjects s "
            sql &= "on s.name = x.name "
            sql &= "and s.parent_obj = x.id "
            sql &= "and s.xtype = 'PK' "
            sql &= "where x.id = object_id('dbo.ppp') "
            sql &= "and x.name not like '_WA_%' "
            sql &= "order by 2, 1, 3"
        Else
            sql = "select i.name"
            sql &= ",ic.index_column_id"
            sql &= ",c.name ColumnName"
            sql &= ",ic.key_ordinal"
            sql &= ",ic.partition_ordinal"
            sql &= ",i.type"
            sql &= ",i.is_primary_key"
            sql &= ",i.is_unique"
            sql &= ",ic.is_descending_key"
            sql &= ",ic.is_included_column"
            sql &= ",d.name filegroup"
            sql &= ",i.fill_factor FILL_FACTOR"
            sql &= ",i.is_padded PAD_INDEX"
            sql &= ",i.ignore_dup_key IGNORE_DUP_KEY"
            sql &= ",i.allow_row_locks ALLOW_ROW_LOCKS"
            sql &= ",i.allow_page_locks ALLOW_PAGE_LOCKS"
            sql &= ",s.no_recompute "
            sql &= "from sys.indexes i "
            sql &= "join sys.stats s "
            sql &= "on i.object_id = s.object_id "
            sql &= "and i.index_id = s.stats_id "
            sql &= "join sys.index_columns ic "
            sql &= "on ic.object_id = i.object_id "
            sql &= "and ic.index_id = i.index_id "
            sql &= "join sys.columns c "
            sql &= "on c.object_id = i.object_id "
            sql &= "and c.column_id = ic.column_id "
            sql &= "join sys.data_spaces d "
            sql &= "on d.data_space_id = i.data_space_id "
            sql &= "where i.object_id = object_id('" & Schema & "." & TableName & "') "
            sql &= "order by 1, 2"
        End If

        Return GetTable(sql)
    End Function

    Public Function TableCheckConstraints(ByVal TableName As String, ByVal Schema As String) As DataTable
        Dim sql As String

        openConnect()

        If Version < 90 Then            'SQL 2000 compatible
            sql = "declare @z bit set @z = 0 "
            sql &= "select c.name ConstraintName"
            sql &= ",m.text Definition"
            sql &= ",@z Replicated"
            sql &= ",@z SystemName"
            sql &= ",null ColumnName"
            sql &= ",@z IsSystem "
            sql &= "from dbo.sysobjects c "
            sql &= "join dbo.syscomments m "
            sql &= "on m.id = c.id "
            sql &= "where c.xtype = 'C' "
            sql &= "and c.parent_obj = object_id('" & Schema & "." & TableName & "') "
            sql &= "order by 1"
        Else
            sql = "select c.name ConstraintName"
            sql &= ",c.definition Definition"
            sql &= ",c.is_not_for_replication Replicated"
            sql &= ",c.is_system_named SystemName"
            sql &= ",l.name ColumnName"
            sql &= ",c.is_ms_shipped IsSystem "
            sql &= "from sys.check_constraints c "
            sql &= "left join sys.columns l "
            sql &= "on l.object_id=c.parent_object_id "
            sql &= "and l.column_id=parent_column_id "
            sql &= "where c.parent_object_id = object_id('" & Schema & "." & TableName & "') "
            sql &= "order by 1"
        End If

        Return GetTable(sql)
    End Function

    Public Function TableTriggers(ByVal TableName As String, ByVal Schema As String) As DataTable
        Dim sql As String

        openConnect()
        sql = "select o.name TriggerName "
        sql &= "from dbo.sysobjects o "
        sql &= "where o.type = 'TR' "
        sql &= "and o.parent_obj = object_id('" & Schema & "." & TableName & "')"

        Return GetTable(sql)
    End Function

    Public Function TablePermissions(ByVal TableName As String, ByVal Schema As String) As DataTable
        Dim sql As String
        openConnect()

        If Version < 90 Then            'SQL 2000 compatible
            sql = "select user_name(p.uid) grantee,x.permission_name"
            sql &= ",case p.protecttype when 204 then 'GRANT_WITH_GRANT_OPTION' when 205 then 'GRANT' when 206 then 'DENY' else '' end state"
            sql &= ",cast(p.columns as integer) columns from dbo.sysprotects p "
            sql &= "join (select 193 id,'SELECT' permission_name union select 195, 'INSERT' union select 196, 'DELETE' union select 197, 'UPDATE') x "
            sql &= "on x.id = p.action where p.id=object_id('" & schema & "." & TableName & "')"
        Else
            sql = "select user_name(grantee_principal_id) grantee,permission_name,state_desc state,sum(power(2, minor_id)) columns "
            sql &= "from sys.database_permissions "
            sql &= "where  permission_name in ('SELECT','INSERT','DELETE','UPDATE','REFERENCES') "
            sql &= "and major_id=object_id('" & schema & "." & TableName & "') "
            sql &= "group by object_name(major_id),user_name(grantee_principal_id),permission_name,state_desc"
        End If

        Return GetTable(sql)
    End Function

    Public Function FunctionColumns(ByVal Name As String, ByVal Schema As String) As DataTable
        Dim sql As String

        openConnect()

        If Version < 90 Then            'SQL 2000 compatible
            sql = "select name from dbo.syscolumns where id = object_id('" & Schema & "." & Name & "') "
            sql &= "and substring(name,1,1) <> '@' order by colid"
        Else
            sql = "select name from sys.columns where object_id = object_id('" & Schema & "." & Name & "') "
            sql &= "and substring(name,1,1) <> '@' order by column_id"
        End If

        Return GetTable(sql)
    End Function

    Public Function TableData(ByVal Name As String, ByVal Schema As String, ByVal Filter As String) As DataTable
        Dim sql As String

        openConnect()
        sql = "select * from " & Schema & "." & Me.QuoteIdentifier(Name)
        If Filter <> "" Then
            sql &= " where " & Filter
        End If

        Return GetTable(sql)
    End Function

    Public Function ProcParms(ByVal Name As String, ByVal Schema As String) As DataTable
        Dim sql As String

        openConnect()
        sql = "select SPECIFIC_NAME"
        sql &= ",ORDINAL_POSITION"
        sql &= ",PARAMETER_NAME"
        sql &= ",DATA_TYPE"
        sql &= ",CHARACTER_MAXIMUM_LENGTH"
        sql &= ",NUMERIC_PRECISION"
        sql &= ",NUMERIC_SCALE"
        sql &= ",PARAMETER_MODE "
        sql &= "from INFORMATION_SCHEMA.PARAMETERS "
        sql &= "where SPECIFIC_NAME='" & Name & "' "
        sql &= "and SPECIFIC_SCHEMA='" & Schema & "' "
        sql &= "order by ORDINAL_POSITION"

        Return GetTable(sql)
    End Function

    Public Function ProcPermissions(ByVal Name As String, ByVal Schema As String) As DataTable
        Dim sql As String

        openConnect()

        If Version < 90 Then            'SQL 2000 compatible
            sql = "select user_name(p.uid) grantee,case p.protecttype "
            sql &= "when 204 then 'W' when 205 then 'G' when 206 then 'D' else '' end state "
            sql &= "from dbo.sysprotects p "
            sql &= "where p.action = 224 and p.id=object_id('" & Schema & "." & Name & "')"
        Else
            sql = "select user_name(grantee_principal_id) grantee,state "
            sql &= "from sys.database_permissions "
            sql &= "where type = 'EX' and major_id=object_id('" & Schema & "." & Name & "')"
        End If
        Return GetTable(sql)
    End Function

    Public Function JobList(ByVal Name As String) As DataTable
        Dim sql As String

        openConnect()
        sql = "select job_id from dbo.sysjobs"
        If Name <> "" Then
            sql &= " where name like '" & Name & "'"
        End If
        sql &= " order by name"

        Return GetTable(sql)
    End Function

    Public Function JobObject(ByVal ID As String) As DataRow
        Dim sql As String

        openConnect()

        If Version < 90 Then            'SQL 2000 compatible
            sql = "select j.name,j.notify_level_eventlog,j.notify_level_email,"
            sql &= "j.notify_level_netsend,j.notify_level_page,j.delete_level,"
            sql &= "o1.name email,o2.name netsend,o3.name page,j.description,"
            sql &= "c.name category,suser_sname(j.owner_sid) owner "
            sql &= "from dbo.sysjobs j "
            sql &= "join dbo.syscategories c "
            sql &= "on c.category_id=j.category_id "
            sql &= "left join dbo.sysoperators o1 "
            sql &= "on o1.id = j.notify_email_operator_id "
            sql &= "left join dbo.sysoperators o2 "
            sql &= "on o2.id = j.notify_netsend_operator_id "
            sql &= "left join dbo.sysoperators o3 "
            sql &= "on o3.id = j.notify_page_operator_id "
            sql &= "where job_id='" & ID & "'"
        Else
            sql = "select j.name,j.notify_level_eventlog,j.notify_level_email,"
            sql &= "j.notify_level_netsend,j.notify_level_page,j.delete_level,"
            sql &= "o1.name email,o2.name netsend,o3.name page,j.description,"
            sql &= "c.name category,suser_sname(j.owner_sid) owner "
            sql &= "from dbo.sysjobs j "
            sql &= "join dbo.syscategories c "
            sql &= "on c.category_id=j.category_id "
            sql &= "left join dbo.sysoperators o1 "
            sql &= "on o1.id = j.notify_email_operator_id "
            sql &= "left join dbo.sysoperators o2 "
            sql &= "on o2.id = j.notify_netsend_operator_id "
            sql &= "left join dbo.sysoperators o3 "
            sql &= "on o3.id = j.notify_page_operator_id "
            sql &= "where job_id='" & ID & "'"
        End If
        Return GetRow(sql)
    End Function

    Public Function JobServer(ByVal ID As String) As DataRow
        Dim sql As String

        openConnect()
        If Version < 90 Then            'SQL 2000 compatible
            sql = "select j.server_id,s.srvname server "
            sql &= "from dbo.sysjobservers j "
            sql &= "join master.dbo.sysservers s on j.server_id=s.srvid "
            sql &= "where job_id='" & ID & "'"
        Else
            sql = "select j.server_id,s.name server "
            sql &= "from dbo.sysjobservers j "
            sql &= "join master.sys.servers s on j.server_id=s.server_id "
            sql &= "where job_id='" & ID & "'"
        End If
        Return GetRow(sql)
    End Function

    Public Function JobStep(ByVal ID As String) As DataTable
        Dim sql As String

        openConnect()

        If Version < 90 Then            'SQL 2000 compatible
            sql = "select j.step_id,j.step_name,j.subsystem,j.command,j.additional_parameters,"
            sql &= "j.cmdexec_success_code,j.on_success_action,j.on_success_step_id,"
            sql &= "j.on_fail_action,j.on_fail_step_id,j.server,j.database_name,"
            sql &= "j.database_user_name,j.retry_attempts,j.retry_interval,j.os_run_priority,"
            sql &= "j.output_file_name,j.flags,null proxy "
            sql &= "from dbo.sysjobsteps j where j.job_id='" & ID & "'"
        Else
            sql = "select j.step_id,j.step_name,j.subsystem,j.command,j.additional_parameters,"
            sql &= "j.cmdexec_success_code,j.on_success_action,j.on_success_step_id,"
            sql &= "j.on_fail_action,j.on_fail_step_id,j.server,j.database_name,"
            sql &= "j.database_user_name,j.retry_attempts,j.retry_interval,j.os_run_priority,"
            sql &= "j.output_file_name,j.flags,p.name proxy "
            sql &= "from dbo.sysjobsteps j left join dbo.sysproxies p on p.proxy_id = j.proxy_id "
            sql &= "where j.job_id='" & ID & "'"
        End If

        Return GetTable(sql)
    End Function

    Public Function JobSchedule(ByVal ID As String) As DataTable
        Dim sql As String

        openConnect()

        If Version < 90 Then            'SQL 2000 compatible
            sql = "select j.name,j.freq_type,j.freq_interval,j.freq_subday_type,"
            sql &= "j.freq_subday_interval,j.freq_relative_interval,j.freq_recurrence_factor,"
            sql &= "j.active_end_date, j.active_start_time, j.active_end_time "
            sql &= "from dbo.sysjobschedules j "
            sql &= "where j.job_id='" & ID & "' and j.enabled=1"
        Else
            sql = "select s.name,s.freq_type,s.freq_interval,s.freq_subday_type,"
            sql &= "s.freq_subday_interval,s.freq_relative_interval,s.freq_recurrence_factor,"
            sql &= "s.active_end_date, s.active_start_time, s.active_end_time "
            sql &= "from dbo.sysjobschedules j "
            sql &= "join dbo.sysschedules s on s.Schedule_id = j.schedule_id "
            sql &= "where j.job_id='" & ID & "' and s.enabled=1"
        End If
        Return GetTable(sql)
    End Function

    Public Function UserCreate(ByVal Logon As String) As String
        Dim s As String
        Dim sql As String

        openConnect()
        If Version < 90 Then            'SQL 2000 compatible
            s = "dbo.sysusers"
        Else
            s = "sys.sysusers"
        End If

        sql = "if not exists" & vbCrLf
        sql &= "(" & vbCrLf
        sql &= "    select  'a'" & vbCrLf
        sql &= "    from    " & s & vbCrLf
        sql &= "    where   name = '" & Logon & "'" & vbCrLf
        sql &= ")" & vbCrLf
        sql &= "begin" & vbCrLf
        If Version < 90 Then            'SQL 2000 compatible
            sql &= "    execute dbo.sp_grantdbaccess @loginame='" & Logon & "',  @name_in_db = '" & Logon & "'" & vbCrLf
        Else
            sql &= "    create User " & Logon & " for login " & Logon & vbCrLf
        End If
        sql &= "end"

        Return sql
    End Function

    Public Function UserGrant(ByVal Database As String, ByVal Logon As String) As String
        Dim sql As String = ""

        openConnect()
        Select Case Database
            Case "master"
                If Version < 90 Then            'SQL 2000 compatible
                    sql = ""
                Else
                    sql = "grant view any definition to " & Logon
                End If

            Case "msdb"
                If Version < 90 Then            'SQL 2000 compatible
                    sql = "grant select on dbo.sysjobs to " & Logon & vbCrLf
                    sql &= "grant select on dbo.sysoperators to " & Logon & vbCrLf
                    sql &= "grant select on dbo.syscategories to " & Logon & vbCrLf
                    sql &= "grant select on dbo.sysjobservers to " & Logon & vbCrLf
                    sql &= "grant select on dbo.sysjobsteps to " & Logon & vbCrLf
                    sql &= "grant select on dbo.sysjobschedules to " & Logon
                Else
                    sql = "grant select on dbo.sysjobs to " & Logon & vbCrLf
                    sql &= "grant select on dbo.sysoperators to " & Logon & vbCrLf
                    sql &= "grant select on dbo.syscategories to " & Logon & vbCrLf
                    sql &= "grant select on dbo.sysjobservers to " & Logon & vbCrLf
                    sql &= "grant select on dbo.sysproxies to " & Logon & vbCrLf
                    sql &= "grant select on dbo.sysjobsteps to " & Logon & vbCrLf
                    sql &= "grant select on dbo.sysschedules to " & Logon & vbCrLf
                    sql &= "grant select on dbo.sysjobschedules to " & Logon
                End If

            Case Else
                If Version < 90 Then            'SQL 2000 compatible
                    sql = "grant select on sysobjects to " & Logon & vbCrLf
                    sql &= "grant select on syscomments to " & Logon & vbCrLf
                    sql &= "grant select on syscolumns to " & Logon & vbCrLf
                    sql &= "grant select on systypes to " & Logon & vbCrLf
                    sql &= "grant select on sysindexes to " & Logon & vbCrLf
                    sql &= "grant select on sysindexkeys to " & Logon & vbCrLf
                    sql &= "grant select on sysforeignkeys to " & Logon
                Else
                    sql &= "grant select on sys.syscomments to " & Logon
                End If

        End Select
        Return sql
    End Function

    Public Function getName(ByVal sText As String) As String
        Dim sOut As String
        Dim Posn As Integer = 1

        sOut = GetNextToken(sText, Posn)
        Select Case LCase(sOut)
            Case "create", "alter"
                sOut = LCase(GetNextToken(sText, Posn))
                If sOut = "procedure" Or sOut = "proc" _
                Or sOut = "function" Or sOut = "func" _
                Or sOut = "view" Or sOut = "trigger" Then
                    sOut = GetNextToken(sText, Posn)
                Else
                    sOut = ""
                End If

            Case Else
                sOut = ""
        End Select
        Return RemoveSquares(sOut)
    End Function

    Public Function ShellName() As String
        Dim sql As String
        Dim s As String = ""
        Dim dr As DataRow

        openConnect()
        sql = "select dbo.shlVariableGet('SystemName')"
        dr = GetRow(sql)
        If Not dr Is Nothing Then
            s = GetString(dr.Item(0))
        End If
        Return s
    End Function

    Public Function GetString(ByVal objValue As Object) As String
        If IsDBNull(objValue) Then
            Return ""
        ElseIf objValue Is Nothing Then
            Return ""
        Else
            Try
                Return CType(objValue, String).TrimEnd
            Catch ex As Exception
                Return objValue.ToString
            End Try
        End If
    End Function

    Public Function GetInteger(ByVal objValue As Object, ByVal iDefault As Integer) As Integer
        If IsDBNull(objValue) Then
            Return iDefault
        ElseIf objValue Is Nothing Then
            Return iDefault
        Else
            Try
                Return CInt(objValue)
            Catch ex As Exception
                Return iDefault
            End Try
        End If
    End Function

    Public Function GetBit(ByVal objValue As Object, ByVal bDefault As Boolean) As Boolean
        If IsDBNull(objValue) Then
            Return bDefault
        ElseIf objValue Is Nothing Then
            Return bDefault
        Else
            Try
                Return CType(objValue, Boolean)
            Catch ex As Exception
                Return bDefault
            End Try
        End If
    End Function

    Public Function GetSQLString(ByVal objValue As Object) As String
        Return Replace(GetString(objValue), "'", "''")
    End Function

    Public Function QuoteIdentifier(ByVal objValue As Object) As String
        Dim s As String = GetString(objValue)
        Dim ss As String
        Dim b As Boolean
        Dim i As Integer

        Select Case LCase(s)
            Case "encryption", "order", "add", "end", "outer", "all", "errlvl", "over", "alter", _
                 "escape", "percent", "and", "except", "plan", "any", "exec", "precision", "as", _
                 "execute", "primary", "asc", "exists", "print", "authorization", "exit", "proc", _
                 "avg", "expression", "procedure", "backup", "fetch", "public", "begin", "file", _
                 "raiserror", "between", "fillfactor", "read", "break", "for", "readtext", _
                 "browse", "foreign", "reconfigure", "bulk", "freetext", "references", "by", _
                 "freetexttable", "replication", "cascade", "from", "restore", "case", "full", _
                 "restrict", "check", "function", "return", "checkpoint", "goto", "revoke", _
                 "close", "grant", "right", "clustered", "group", "rollback", "coalesce", _
                 "having", "rowcount", "collate", "holdlock", "rowguidcol", "column", "identity", _
                 "rule", "commit", "identity_insert", "save", "compute", "identitycol", "schema", _
                 "constraint", "if", "select", "contains", "in", "session_user", "containstable", _
                 "index", "set", "continue", "inner", "setuser", "convert", "insert", "shutdown", _
                 "count", "intersect", "some", "create", "into", "statistics", "cross", "is", _
                 "sum", "current", "join", "system_user", "current_date", "key", "table", _
                 "current_time", "kill", "textsize", "current_timestamp", "left", "then", _
                 "current_user", "like", "to", "cursor", "lineno", "top", "database", "load", _
                 "tran", "databasepassword", "max", "transaction", "dateadd", "min", "trigger", _
                 "datediff", "national", "truncate", "datename", "nocheck", "tsequal", "datepart", _
                 "nonclustered", "union", "dbcc", "not", "unique", "deallocate", "null", "update", _
                 "declare", "nullif", "updatetext", "default", "of", "use", "delete", "off", _
                 "user", "deny", "offsets", "values", "desc", "on", "varying", "disk", "open", _
                 "view", "distinct", "opendatasource", "waitfor", "distributed", "openquery", _
                 "when", "double", "openrowset", "where", "drop", "openxml", "while", "dump", _
                 "option", "with", "else", "or", "writetext"
                s = """" & s & """"
            Case Else
                b = False
                For i = 1 To Len(s)
                    ss = LCase(Mid(s, i, 1))
                    If InStr("abcdefghijklmnopqrstuvwxyz@_$#0123456789", ss, CompareMethod.Text) = 0 Then
                        b = True
                        Exit For
                    End If
                Next
                If b Then
                    s = Replace(s, """", """""")
                    s = """" & s & """"
                End If

                'The first character must be one of the following: 
                'A letter as defined by the Unicode Standard 3.2. The Unicode definition of letters includes Latin characters from a through z, from A through Z, and also letter characters from other languages.
                'The _, @, or #. 
                'Certain symbols at the beginning of an identifier have special meaning in SQL Server. A regular identifier that starts with the at sign always denotes a local variable or parameter and cannot be used as the name of any other type of object. An identifier that starts with a number sign denotes a temporary table or procedure. An identifier that starts with double number signs (##) denotes a global temporary object. Although the number sign or double number sign characters can be used to begin the names of other types of objects, we do not recommend this practice.
                'Some Transact-SQL functions have names that start with double at signs (@@). To avoid confusion with these functions, you should not use names that start with @@. 

                'Subsequent characters can include the following: 
                'Letters as defined in the Unicode Standard 3.2.
                'Decimal numbers from either Basic Latin or other national scripts.
                'The @, $, #, or _.
                'The identifier must not be a Transact-SQL reserved word. 
                'SQL Server reserves both the uppercase and lowercase versions of reserved words.
                'Embedded spaces or special characters are not allowed.
                'Supplementary characters are not allowed.


        End Select

        Return s
    End Function

    Public Function CleanConstraint(ByVal constratint As String) As String
        Dim s As String

        s = RemoveSquares(constratint)
        If Mid(s, 1, 1) = "(" And Right(s, 1) = ")" Then
            s = Mid(s, 2, Len(s) - 2)
        End If

        Return s
    End Function

    Public Function RemoveSquares(ByVal sText As String) As String
        Dim s As String = ""
        Dim ss As String
        Dim sSave As String = ""
        Dim i As Integer
        Dim mode As Integer = 0
        Dim bc As Integer = 0

        For i = 1 To Len(sText)
            ss = Mid(sText, i, 1)
            Select Case mode
                Case 0
                    Select Case ss
                        Case "["
                            bc = 1
                            sSave = ""
                            mode = 3

                        Case "'"
                            mode = 1
                            s &= ss

                        Case """"
                            sSave = ""
                            mode = 2

                        Case Else
                            s &= ss

                    End Select

                Case 1
                    If ss = "'" Then mode = 0
                    s &= ss

                Case 2
                    If ss = """" Then
                        If Mid(sText, i + 1, 1) = """" Then
                            sSave &= """"
                            i += 1
                        Else
                            s &= QuoteIdentifier(sSave)
                            mode = 0
                        End If
                    Else
                        sSave &= ss
                    End If

                Case 3
                    Select Case ss
                        Case "]"
                            bc -= 1
                            If bc = 0 Then
                                s &= QuoteIdentifier(sSave)
                                mode = 0
                            Else
                                sSave &= ss
                            End If

                        Case "["
                            sSave &= ss
                            bc += 1

                        Case Else
                            sSave &= ss

                    End Select
            End Select
        Next
        Return s
    End Function
#End Region

#Region "Private functions"

    Private Function getConnectString() As String
        Dim con As String

        If sConnect <> "" Then
            Return sConnect
        End If

        con = "server=" & sServer & ";database=" & sDatabase
        If sNetwork <> "" Then con &= ";Network=" & sNetwork
        If sUserID <> "" Then
            con &= ";User ID=" & sUserID
            If sPassword <> "" Then
                con &= ";pwd=" & sPassword
            End If
        Else
            con &= ";Integrated Security=SSPI"
        End If
        Return con
    End Function

    Private Function GetTable(ByVal sql As String) As DataTable
        Dim psAdapt As SqlDataAdapter
        Dim Details As New DataSet

        psAdapt = New SqlDataAdapter(sql, psConn)
        If iTimeOut > 0 Then
            psAdapt.SelectCommand.CommandTimeout = iTimeOut
        End If
        psAdapt.SelectCommand.CommandType = CommandType.Text

        psAdapt.Fill(Details)

        If Details.Tables.Count > 0 Then
            Return Details.Tables(0)
        End If

        Return Nothing
    End Function

    Private Function GetRow(ByVal sql As String) As DataRow
        Dim dt As DataTable

        dt = GetTable(sql)
        If Not dt Is Nothing Then
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            End If
        End If
        Return Nothing
    End Function

    Private Sub openConnect()
        Dim sql As String
        Dim i As Integer = -1
        Dim dt As DataTable

        If Connected Then
            Return
        End If

        psConn = New SqlConnection(getConnectString())
        psConn.Open()
        Connected = True

        sql = "select cmptlevel from master.dbo.sysdatabases where name = db_name()"
        dt = GetTable(sql)
        If Not dt Is Nothing Then
            Version = CInt(dt.Rows(0).Item(0))
        End If
    End Sub

    Private Sub CloseConnect()
        If Connected Then
            psConn.Close()
            Version = -1
            Connected = False
        End If
    End Sub

    Private Function GetNextToken(ByVal sSQL As String, ByRef Start As Integer) As String
        Dim ls As String = ""
        Dim Mode As Integer = 0

        ' liMode =
        ' 0 = searching for tokening
        ' 1 = line comment
        ' 2 = getting parameter no quotes
        ' 3 = getting parameter single quote
        ' 4 = getting parameter double quote
        ' 5 = block comment

        Do While (Start <= Len(sSQL))
            Select Case Mode
                Case 0
                    Select Case Mid(sSQL, Start, 1)
                        Case " ", Chr(9), Chr(10), Chr(13), "=", ","
                        Case "-"
                            If Mid(sSQL, Start, 2) = "--" Then
                                Start += 1
                                Mode = 1
                            Else
                                Mode = 2
                                ls = Mid(sSQL, Start, 1)
                            End If
                        Case "/"
                            If Mid(sSQL, Start, 2) = "/*" Then
                                Start += 1
                                Mode = 5
                            Else
                                Mode = 2
                                ls = Mid(sSQL, Start, 1)
                            End If
                        Case "'"
                            ls = Mid(sSQL, Start, 1)
                            Mode = 3
                        Case Chr(34)
                            ls = Mid(sSQL, Start, 1)
                            Mode = 4
                        Case Else
                            Mode = 2
                            ls = Mid(sSQL, Start, 1)
                    End Select
                Case 1              ' 1 = line comment
                    If Mid(sSQL, Start, 1) = Chr(10) Then
                        Mode = 0
                    End If
                Case 2              ' 2 = getting parameter no quotes
                    Select Case Mid(sSQL, Start, 1)
                        Case " ", Chr(9), Chr(10), Chr(13), ",", "("
                            Exit Do
                        Case Else
                            ls &= Mid(sSQL, Start, 1)
                    End Select
                Case 3              ' 3 = getting parameter single quote
                    ls &= Mid(sSQL, Start, 1)
                    If Mid$(sSQL, Start, 1) = "'" Then
                        Exit Do
                    End If
                Case 4              ' 4 = getting parameter double quote
                    ls &= Mid(sSQL, Start, 1)
                    If Mid(sSQL, Start, 1) = Chr(34) Then
                        If Mid(sSQL, Start + 1, 1) = Chr(34) Then
                            Start += 1
                        Else
                            Exit Do
                        End If
                    End If
                Case 5              ' 5 = block comment
                    If Mid(sSQL, Start, 2) = "*/" Then
                        Start += 1
                        Mode = 0
                    End If
            End Select
            Start += 1
        Loop
        GetNextToken = ls
    End Function
#End Region
End Class

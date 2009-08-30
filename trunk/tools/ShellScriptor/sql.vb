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

    Public Version As Integer = -1
#End Region

#Region "Methods"
    Public Function DatabaseList() As DataTable
        Dim sql As String

        openConnect()
        sql = "select name,cmptlevel from master.dbo.sysdatabases where name not in ('master','tempdb','model')"

        Return GetTable(sql)
    End Function

    Public Function DatabaseObject(ByVal Name As String, ByVal Type As String) As DataTable
        Dim sql As String

        openConnect()
        sql = "select type, name "
        sql &= "from dbo.sysobjects "
        Select Case Type
            Case "P", "U", "V"
                sql &= "where type = '" & Type & "' "
            Case "F"
                sql &= "where type in ('FN', 'TF') "
            Case Else
                sql &= "where type in ('P', 'U', 'V', 'FN', 'TF') "
        End Select
        If Name <> "" Then
            sql &= "and name like '" & Name & "' "
        End If
        sql &= "and uid = 1 "
        sql &= "and name not like 'dt_%' "
        sql &= "and name not in ('syssegments','sysconstraints','sysdiagrams',"
        sql &= "'sp_alterdiagram','sp_creatediagram',"
        sql &= "'sp_dropdiagram','sp_helpdiagramdefinition','sp_helpdiagrams',"
        sql &= "'sp_renamediagram','sp_upgraddiagrams','fn_diagramobjects') "
        sql &= "order by type, name"

        Return GetTable(sql)
    End Function

    Public Function ObjectText(ByVal Name As String) As DataTable
        Dim sql As String

        openConnect()
        sql = "select text from syscomments"
        sql &= " where id = object_id('" & Name & "') order by number, colid"

        Return GetTable(sql)
    End Function

    Public Function ObjectSettings(ByVal Name As String) As DataTable
        Dim sql As String

        openConnect()
        sql = "select uses_ansi_nulls nulls,uses_quoted_identifier quoted from sys.sql_modules"
        sql &= " where object_id=object_id('" & Name & "')"

        Return GetTable(sql)
    End Function

    Public Function TableColumns(ByVal Name As String) As DataTable
        Dim sql As String

        openConnect()
        sql = "select TABLE_NAME TableName"
        sql &= ",i.COLUMN_NAME"
        sql &= ",i.DATA_TYPE"
        sql &= ",i.CHARACTER_MAXIMUM_LENGTH"
        sql &= ",i.IS_NULLABLE"
        sql &= ",i.NUMERIC_PRECISION"
        sql &= ",i.NUMERIC_SCALE"
        sql &= ",s.name DEFAULT_NAME"
        sql &= ",i.COLUMN_DEFAULT DEFAULT_TEXT "
        sql &= "from INFORMATION_SCHEMA.COLUMNS i "
        sql &= "join dbo.syscolumns c "
        sql &= "on c.id = object_id('" & Name & "') "
        sql &= "and c.name = i.COLUMN_NAME "
        sql &= "left join dbo.sysobjects s "
        sql &= "on s.id = c.cdefault "
        sql &= "where TABLE_NAME = '" & Name & "' "
        sql &= "and TABLE_SCHEMA = 'dbo' "
        sql &= "order by ORDINAL_POSITION"

        Return GetTable(sql)
    End Function

    Public Function TableIdentity(ByVal Name As String) As String
        Dim sql As String
        Dim dr As DataRow

        openConnect()
        sql = "select name from syscolumns where id = object_id('" & Name & "') and colstat & 1 = 1"
        dr = GetRow(sql)
        If Not dr Is Nothing Then
            Return GetString(dr.Item("name"))
        End If
        Return ""
    End Function

    Public Function TableFKeys(ByVal Name As String) As DataTable
        Dim sql As String

        openConnect()
        sql = "select c.CONSTRAINT_NAME ConstraintName"
        sql &= ",u1.ORDINAL_POSITION Sequence"
        sql &= ",u1.COLUMN_NAME ColumnName"
        sql &= ",u2.TABLE_NAME LinkedTable"
        sql &= ",u2.COLUMN_NAME LinkedColumn "
        sql &= "from INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS c "
        sql &= "join INFORMATION_SCHEMA.KEY_COLUMN_USAGE u1 "
        sql &= "on u1.CONSTRAINT_CATALOG = c.CONSTRAINT_CATALOG "
        sql &= "and u1.CONSTRAINT_SCHEMA = c.CONSTRAINT_SCHEMA "
        sql &= "and u1.CONSTRAINT_NAME = c.CONSTRAINT_NAME "
        sql &= "join INFORMATION_SCHEMA.KEY_COLUMN_USAGE u2 "
        sql &= "on u2.CONSTRAINT_CATALOG = c.UNIQUE_CONSTRAINT_CATALOG "
        sql &= "and u2.CONSTRAINT_SCHEMA = c.UNIQUE_CONSTRAINT_SCHEMA "
        sql &= "and u2.CONSTRAINT_NAME = c.UNIQUE_CONSTRAINT_NAME "
        sql &= "and u2.ORDINAL_POSITION = u1.ORDINAL_POSITION "
        sql &= "where u1.TABLE_NAME = '" & Name & "' "
        sql &= "order by 1, 2"

        Return GetTable(sql)
    End Function

    Public Function TableIndexes(ByVal Name As String) As DataTable
        Dim sql As String

        openConnect()

        If Version < 90 Then            'SQL 2000 compatible
            sql = "select i.keyno key_ordinal"
            sql &= ",x.name"
            sql &= ",index_col(object_name(x.id), x.indid, i.keyno) ColumnName"
            sql &= ",case indexkey_property(x.id, x.indid, i.colid, 'isdescending') when 1 then 1 else 0 end is_descending_key"
            sql &= ",case when indexproperty(x.id, x.name, 'IsClustered') = 1 then 1 else 2 end type"
            sql &= ",case when s.name is not null then 1 else 0 end is_primary_key"
            sql &= ",case when indexproperty(x.id, x.name, 'IsUnique') = 1 then 1 else 0 end is_unique"
            sql &= ",0 is_included_column "
            sql &= "from dbo.sysindexes x "
            sql &= "join dbo.sysindexkeys i "
            sql &= "on i.id = x.id "
            sql &= "and i.indid = x.indid "
            sql &= "left join dbo.sysobjects s "
            sql &= "on s.name = x.name "
            sql &= "and s.parent_obj = x.id "
            sql &= "and s.xtype = 'PK' "
            sql &= "where x.id = object_id('" & Name & "') "
            sql &= "and x.name not like '_WA_%' "
            sql &= "order by 2, 1, 3"
        Else
            sql = "select ic.key_ordinal,i.name,c.name ColumnName,"
            sql &= "ic.is_descending_key,i.type,"
            sql &= "i.is_primary_key,"
            sql &= "i.is_unique,ic.is_included_column "
            sql &= "from sys.indexes i "
            sql &= "join sys.index_columns ic "
            sql &= "on ic.object_id = i.object_id "
            sql &= "and ic.index_id = i.index_id "
            sql &= "join sys.columns c "
            sql &= "on c.object_id = i.object_id "
            sql &= "and c.column_id = ic.column_id "
            sql &= "where i.object_id = object_id('" & Name & "') "
            sql &= "order by 2, 1, 3"
        End If

        Return GetTable(sql)
    End Function

    Public Function TableData(ByVal Name As String, ByVal Filter As String) As DataTable
        Dim sql As String

        openConnect()
        sql = "select * from dbo." & Name
        If Filter <> "" Then
            sql &= " where " & Filter
        End If

        Return GetTable(sql)
    End Function

    Public Function TableTriggers(ByVal Name As String) As DataTable
        Dim sql As String

        openConnect()
        sql = "select o.name TriggerName "
        sql &= "from dbo.sysobjects o "
        sql &= "where o.type = 'TR' "
        sql &= "and o.parent_obj = object_id('" & Name & "')"

        Return GetTable(sql)
    End Function

    Public Function ProcParms(ByVal Name As String) As DataTable
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
        sql &= "where SPECIFIC_NAME='" & Name
        sql &= "' order by ORDINAL_POSITION"

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
        sql = "select j.*,suser_sname(j.owner_sid) owner,c.name category,"
        sql &= "o1.name email,o2.name netsend,o3.name page "
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

        Return GetRow(sql)
    End Function

    Public Function JobServer(ByVal ID As String) As DataRow
        Dim sql As String

        openConnect()
        sql = "select j.*,s.name server "
        sql &= "from dbo.sysjobservers j "
        sql &= "join master.sys.servers s on j.server_id=s.server_id "
        sql &= "where job_id='" & ID & "'"

        Return GetRow(sql)
    End Function

    Public Function JobStep(ByVal ID As String) As DataTable
        Dim sql As String

        openConnect()
        sql = "select j.*,p.name proxy from dbo.sysjobsteps j "
        sql &= "left join dbo.sysproxies p on p.proxy_id = j.proxy_id "
        sql &= "where job_id='" & ID & "'"

        Return GetTable(sql)
    End Function

    Public Function JobSchedule(ByVal ID As String) As DataTable
        Dim sql As String

        openConnect()
        sql = "select * from dbo.sysjobschedules j "
        sql &= "join dbo.sysschedules s on s.Schedule_id = j.schedule_id "
        sql &= "where j.job_id='" & ID & "' and s.enabled=1"

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
        sql &= "    create User " & Logon & " for login " & Logon & vbCrLf
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
                    sql = ""
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
                sql = ""

        End Select
        Return sql
    End Function

    Public Function ShellName() As String
        Dim sql As String
        Dim s As String = ""
        Dim dr As DataRow

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

    Public Function GetSQLString(ByVal objValue As Object) As String
        Return Replace(GetString(objValue), "'", "''")
    End Function

    Public Function QuoteIdentifier(ByVal objValue As Object) As String
        Dim s As String = GetString(objValue)

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
                If InStr(s, " ", CompareMethod.Text) > 0 Then
                    s = """" & s & """"
                End If
        End Select

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
#End Region
End Class

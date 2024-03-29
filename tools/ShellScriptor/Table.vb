Option Explicit On
Option Strict On

'todo
' partition scheme

Imports System.Data.SqlClient

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

Public Class TableColumn
    Private sType As String = ""
    Private iLength As Integer = 0
    Private iPrecision As Integer = 0
    Private iScale As Integer = 0
    Private sCollation As String = ""
    Private bANSIPadded As Boolean = True
    Private sqllib As New sql

#Region "Properties"
    Public Index As Integer
    Public Name As String
    Public Nullable As String
    Public DefaultName As String
    Public DefaultValue As String
    Public Primary As Boolean = False
    Public Descend As Boolean = False
    Public Identity As Boolean = False
    Public RowGuid As Boolean = False
    Public Seed As Integer = 1
    Public Increment As Integer = 1
    Public Replicated As Boolean = False
    Public Computed As String = ""
    Public Persisted As Boolean = False
    Public XMLDocument As Boolean = False
    Public XMLCollection As String = ""

    Public ReadOnly Property QuotedName() As String
        Get
            QuotedName = sqllib.QuoteIdentifier(Name)
        End Get
    End Property

    Public ReadOnly Property QuotedDefaultName() As String
        Get
            QuotedDefaultName = sqllib.QuoteIdentifier(DefaultName)
        End Get
    End Property

    ' bigint
    ' binary(len)
    ' bit
    ' char(len)
    ' datetime
    ' decimal(prec,scale)
    ' float(len)
    ' image
    ' int
    ' money
    ' nchar(len)
    ' ntext
    ' numeric(prec,scale)
    ' nvarchar(len)
    ' real
    ' smalldatetime
    ' smallint
    ' smallmoney
    ' sql_variant
    ' sysname
    ' text
    ' timestamp
    ' tinyint
    ' uniqueidentifier
    ' varbinary(len)
    ' varchar(len)
    'xml

    Public Property Type() As String
        Get
            Type = sType
        End Get
        Set(ByVal value As String)
            If value = "int" Then
                sType = "integer"
            Else
                sType = value
            End If
        End Set
    End Property

    Public Property Length() As Integer
        Get
            Dim i As Integer
            Select Case sType
                Case "bigint", "bit", "datetime", "decimal", _
                    "image", "integer", "money", "ntext", _
                     "numeric", "real", "smalldatetime", _
                    "smallint", "smallmoney", "sql_variant", _
                    "text", "timestamp", "tinyint", _
                    "uniqueidentifier", "xml"
                    i = 0
                Case "sysname"
                    i = 128
                Case Else
                    i = iLength
            End Select
            Length = i
        End Get
        Set(ByVal value As Integer)
            iLength = value
        End Set
    End Property

    Public Property Precision() As Integer
        Get
            Dim i As Integer
            Select Case sType
                Case "decimal", "numeric"
                    i = iPrecision
                Case Else
                    i = 0
            End Select
            Precision = i
        End Get
        Set(ByVal value As Integer)
            iPrecision = value
        End Set
    End Property

    Public Property Scale() As Integer
        Get
            Dim i As Integer
            Select Case sType
                Case "decimal", "numeric"
                    i = iScale
                Case Else
                    i = 0
            End Select
            Scale = i
        End Get
        Set(ByVal value As Integer)
            iScale = value
        End Set
    End Property

    Public Property Collation() As String
        Get
            Dim s As String = ""
            Select Case sType
                Case "text", "ntext", "varchar", "char", "nvarchar", "nchar"
                    s = sCollation
            End Select
            Collation = s
        End Get
        Set(ByVal value As String)
            sCollation = value
        End Set
    End Property

    Public Property ANSIPadded() As String
        Get
            Dim s As String
            Select Case sType
                Case "varchar", "char", "nvarchar", "nchar", _
                     "binary", "varbinary", "sql_variant"
                    If bANSIPadded Then
                        s = "Y"
                    Else
                        s = "N"
                    End If

                Case Else
                    s = "X"
            End Select
            ANSIPadded = s
        End Get
        Set(ByVal value As String)
            If value = "Y" Then
                bANSIPadded = True
            Else
                bANSIPadded = False
            End If
        End Set
    End Property

    Public ReadOnly Property vbType() As String
        Get
            Dim s As String
            Select Case sType
                Case "char", "varchar", "nchar", "nvarchar", "sysname"
                    s = "string"
                Case "decimal", "numeric"
                    s = "double"
                Case "datetime"
                    s = "datetime"
                Case "smalldatetime"
                    s = "date"
                Case "uniqueidentifier"
                    s = "Guid"
                Case Else
                    s = sType
            End Select
            vbType = s
        End Get
    End Property

    Public ReadOnly Property TypeText() As String
        Get
            Dim s As String
            s = sType
            Select Case sType
                Case "varchar", "nvarchar", "varbinary"
                    If iLength = -1 Then
                        s &= "(max)"
                    Else
                        s &= "(" & iLength & ")"
                    End If
                Case "binary", "char", "nchar"
                    s &= "(" & iLength & ")"
                Case "decimal", "numeric"
                    s &= "(" & iPrecision & "," & iScale & ")"
                Case "float"
                    If iPrecision <> 53 Then
                        s &= "(" & iPrecision & ")"
                    End If
                Case "xml"
                    If XMLCollection <> "" Then
                        s &= "("
                        If XMLDocument Then
                            s &= "document "
                        Else
                            s &= "content "
                        End If
                        s &= XMLCollection & ")"
                    End If
            End Select
            TypeText = s
        End Get
    End Property
#End Region

#Region "Methods"
    Public Function DataFormat(ByVal Value As Object) As String
        Dim s As String
        If IsDBNull(Value) Then
            s = "null"
        Else
            Select Case vbType
                Case "string", "Guid"
                    s = "'" & Replace(RTrim(Value.ToString), "'", "''", 1, -1, CompareMethod.Text) & "'"
                Case "datetime"
                    s = "'" & Format(Value, "d-MMM-yyyy hh:mm:ss tt") & "'"
                Case "date"
                    s = "'" & Format(Value, "d-MMM-yyyy") & "'"
                Case Else
                    s = Value.ToString
            End Select
        End If
        Return s
    End Function
#End Region
End Class

Public Class TableColumns
#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New TableColumnsEnum(Keys, Values)
    End Function

    Public Class TableColumnsEnum
        Implements IEnumerable, IEnumerator
        Private Values As New Hashtable
        Dim Keys() As String
        Private EnumeratorPosition As Integer = -1

        Public Sub New(ByVal aKeys() As String, ByVal Hash As Hashtable)
            Keys = aKeys
            Values = Hash
        End Sub

        Public Function GetEnumerator() As System.Collections.IEnumerator _
                            Implements System.Collections.IEnumerable.GetEnumerator
            Return CType(Me, IEnumerator)
        End Function

        Public Overridable Overloads ReadOnly Property Current() As Object _
                                                    Implements IEnumerator.Current
            Get
                Return CType(Values.Item(Keys(EnumeratorPosition)), TableColumn)
            End Get
        End Property

        Public Function MoveNext() As Boolean _
                                Implements System.Collections.IEnumerator.MoveNext
            EnumeratorPosition += 1
            Return (EnumeratorPosition < Values.Count)
        End Function

        Public Overridable Overloads Sub Reset() Implements IEnumerator.Reset
            EnumeratorPosition = -1
        End Sub
    End Class
#End Region

    Private PreLoad As Integer = -1
    Private sTable As String
    Private qTable As String
    Private sSchema As String = "dbo"
    Private qSchema As String = "dbo"
    Private sPKey As String = ""
    Private bPKClust As Boolean = True
    Private sIdentity As String = ""

    Private sFileGroup As String = "PRIMARY"
    Private sTextFileGroup As String = "PRIMARY"
    Private sDefFileGroup As String = "PRIMARY"
    Private sPartitionScheme As String = ""
    Private sSchemeColumn As String = ""
    Private bAudit As Boolean = False
    Private bState As Boolean = False
    Private bConsName As Boolean = True
    Private sDefCollation As String
    Private bCollation As Boolean = False
    Private fixdef As Boolean = False

    Private xPKeys(0) As String
    Private xFKeys(0) As String
    Private xIndexs(0) As String
    'Private xTriggers(0) As String
    Private dtIndexs As DataTable
    Private dtFKeys As DataTable
    Private dtCheck As DataTable
    Private dtPerms As DataTable

    Private Values As New Hashtable
    Private Keys(0) As String
    Private slib As New sql

    ' CREATE TABLE
    '     [ [ database_name . ] [ schema_name . ] table_name
    '         ( { <column_definition> | <computed_column_definition> }
    '         [ <table_constraint> ] [ ,...n ] )
    '     [ ON { filegroup
    'x         | partition_scheme_name ( partition_column_name )
    '-         | "default" } ]
    '     [ { TEXTIMAGE_ON { filegroup | "default" } ]
    '
    ' <column_definition> ::=
    ' column_name <data_type>
    '     [ COLLATE collation_name ]
    '     [ NULL | NOT NULL ]
    '     [
    '         [ IDENTITY [ ( seed , increment ) ]]
    '         [ NOT FOR REPLICATION ]
    '     ]
    '     [ ROWGUIDCOL ]
    '     [ <column_constraint> [ ...n ] ]
    ' 
    ' <data type> ::=
    ' [ type_schema_name . ] type_name
    '     [ ( precision [ , scale ] | max |
    '         [ { CONTENT | DOCUMENT } ] xml_schema_collection ) ]
    '
    ' <column_constraint> ::=
    ' [ CONSTRAINT constraint_name ]
    ' {     { PRIMARY KEY |
    '-        UNIQUE }
    '         [ CLUSTERED | NONCLUSTERED ]
    '         [
    '             WITH ( < index_option > [ , ...n ] )
    '         ]
    '         [ ON { filegroup
    'x            | partition_scheme_name ( partition_column_name )
    '-            | "default" } ]
    '   | [ FOREIGN KEY ]
    '         REFERENCES [ schema_name . ] referenced_table_name [ ( ref_column ) ]
    '         [ ON DELETE { NO ACTION | CASCADE | SET NULL | SET DEFAULT } ]
    '         [ ON UPDATE { NO ACTION | CASCADE | SET NULL | SET DEFAULT } ]
    '         [ NOT FOR REPLICATION ]
    '   | CHECK [ NOT FOR REPLICATION ] ( logical_expression )
    '   | DEFAULT constant_expression
    ' }
    '
    ' <computed_column_definition> ::=
    '  column_name AS computed_column_expression
    '  [ PERSISTED [ NOT NULL ] ]
    '  [
    '     [ CONSTRAINT constraint_name ]
    '     { PRIMARY KEY |
    '-      UNIQUE }
    '         [ CLUSTERED | NONCLUSTERED ]
    '         [
    '             WITH ( <index_option> [ , ...n ] )
    '         ]
    '         [ ON { filegroup
    'x            | partition_scheme_name ( partition_column_name ) | "default" } ]
    '     | [ FOREIGN KEY ]
    '         REFERENCES referenced_table_name [ ( ref_column ) ]
    '         [ ON DELETE { NO ACTION | CASCADE } ]
    '         [ ON UPDATE { NO ACTION } ]
    '         [ NOT FOR REPLICATION ]
    '     | CHECK [ NOT FOR REPLICATION ] ( logical_expression )
    ' ]
    '
    ' < table_constraint > ::=
    ' [ CONSTRAINT constraint_name ]
    ' {
    '     { PRIMARY KEY |
    '-      UNIQUE }
    '         [ CLUSTERED | NONCLUSTERED ]
    '                 (column [ ASC | DESC ] [ ,...n ] )
    '         [
    '             WITH ( <index_option> [ , ...n ] )
    '         ]
    '         [ ON { filegroup
    'x            | partition_scheme_name (partition_column_name) | "default" } ]
    '     | FOREIGN KEY
    '                 ( column [ ,...n ] )
    '         REFERENCES referenced_table_name [ ( ref_column [ ,...n ] ) ]
    '         [ ON DELETE { NO ACTION | CASCADE | SET NULL | SET DEFAULT } ]
    '         [ ON UPDATE { NO ACTION | CASCADE | SET NULL | SET DEFAULT } ]
    '         [ NOT FOR REPLICATION ]
    '     | CHECK [ NOT FOR REPLICATION ] ( logical_expression )
    ' }
    '
    ' <index_option> ::=
    ' {
    '    PAD_INDEX = { ON | OFF }
    '   | FILLFACTOR = fillfactor
    '   | IGNORE_DUP_KEY = { ON | OFF }
    '   | STATISTICS_NORECOMPUTE = { ON | OFF }
    '   | ALLOW_ROW_LOCKS = { ON | OFF}
    '   | ALLOW_PAGE_LOCKS ={ ON | OFF}
    ' }

#Region "Properties"
    Public Property TableName() As String
        Get
            TableName = sTable
        End Get
        Set(ByVal value As String)
            If PreLoad = 0 Then
                sTable = value
                qTable = slib.QuoteIdentifier(sTable)
            End If
        End Set
    End Property

    Public Property Schema() As String
        Get
            Schema = sSchema
        End Get
        Set(ByVal value As String)
            If PreLoad = 0 Then
                sSchema = value
                qSchema = slib.QuoteIdentifier(sSchema)
            End If
        End Set
    End Property

    Public ReadOnly Property PrimaryKey() As String
        Get
            Return sPKey
        End Get
    End Property

    Public ReadOnly Property State() As Integer
        Get
            Return PreLoad
        End Get
    End Property

    Public ReadOnly Property PKeys() As String()
        Get
            Return xPKeys
        End Get
    End Property

    Public ReadOnly Property IKeys() As String()
        Get
            Return xIndexs
        End Get
    End Property

    Public ReadOnly Property FKeys() As String()
        Get
            Return xFKeys
        End Get
    End Property

    Public ReadOnly Property Clustered() As Boolean
        Get
            Return bPKClust
        End Get
    End Property

    Public Property ScriptConstraints() As Boolean
        Get
            ScriptConstraints = bConsName
        End Get
        Set(ByVal value As Boolean)
            bConsName = value
        End Set
    End Property

    Public Property ScriptCollations() As Boolean
        Get
            ScriptCollations = bCollation
        End Get
        Set(ByVal value As Boolean)
            bCollation = value
        End Set
    End Property

    Public ReadOnly Property IdentityColumn() As String
        Get
            Return sIdentity
        End Get
    End Property

    Public ReadOnly Property hasIdentity() As Boolean
        Get
            If sIdentity = "" Then
                Return False
            End If
            Return True
        End Get
    End Property

    Public ReadOnly Property hasAudit() As Boolean
        Get
            Return bAudit
        End Get
    End Property

    Public ReadOnly Property hasState() As Boolean
        Get
            Return bState
        End Get
    End Property

    Public ReadOnly Property Column(ByVal index As String) As TableColumn
        Get
            Try
                Return DirectCast(Values.Item(index), TableColumn)
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property TableText() As String
        Get
            Return CreateTable(False)
        End Get
    End Property

    Public ReadOnly Property FullTableText() As String
        Get
            Return CreateTable(True)
        End Get
    End Property

    Public ReadOnly Property IndexShort(ByVal IndexName As String) As String
        Get
            '  CREATE [ UNIQUE ] [ CLUSTERED | NONCLUSTERED ] INDEX index_name
            '      ON <object> ( column [ ASC | DESC ] [ ,...n ] )
            '      [ INCLUDE ( column_name [ ,...n ] ) ]
            '      [ WITH ( <relational_index_option> [ ,...n ] ) ]
            '      [ ON { filegroup_name
            'x          | partition_scheme_name ( column_name )
            '-          | default
            '           }
            '      ]
            '
            '  <object> ::=
            '  {
            '      [ database_name. [ schema_name ] . | schema_name. ]
            '          table_or_view_name
            '  }
            '
            '  <relational_index_option> ::=
            '  {
            '      PAD_INDEX  = { ON | OFF }
            '    | FILLFACTOR = fillfactor
            '-   | SORT_IN_TEMPDB = { ON | OFF }
            '    | IGNORE_DUP_KEY = { ON | OFF }
            '    | STATISTICS_NORECOMPUTE = { ON | OFF }
            '-   | DROP_EXISTING = { ON | OFF }
            '-   | ONLINE = { ON | OFF }
            '    | ALLOW_ROW_LOCKS = { ON | OFF }
            '    | ALLOW_PAGE_LOCKS = { ON | OFF }
            '-   | MAXDOP = max_degree_of_parallelism
            '  }
            Dim i As Integer = 0
            Dim s As String
            Dim sOut As String = ""
            Dim sInc As String = ""
            Dim sWth As String = ""
            Dim sOn As String = ""
            Dim iClust As Integer

            If dtIndexs Is Nothing Then
                Return ""
            End If

            If dtIndexs.Rows.Count = 0 Then
                Return ""
            End If

            For Each r As DataRow In dtIndexs.Rows
                If CInt(r("is_primary_key")) = 0 Then
                    If IndexName = slib.GetString(r("name")) Then
                        If i = 0 Then
                            iClust = slib.GetInteger(r("type"), 0)
                            sOut &= "create" & slib.GetString(IIf(CInt(r("is_unique")) <> 0, " unique", ""))
                            sOut &= slib.GetString(IIf(iClust = 1, " clustered", " nonclustered"))
                            sOut &= " index " & slib.QuoteIdentifier(IndexName) & " on " & qSchema & "." & qTable & " ("
                            sOut &= slib.QuoteIdentifier(r("ColumnName"))
                            sWth = IndexWith(r)
                            s = slib.GetString(r("filegroup"))
                            If s <> sDefFileGroup And iClust <> 1 Then
                                sOn &= "on " & s & vbCrLf
                            End If
                        Else
                            If CInt(r("is_included_column")) = 0 Then
                                sOut &= "," & slib.QuoteIdentifier(r("ColumnName"))
                            Else
                                sInc &= "," & slib.QuoteIdentifier(r("ColumnName"))
                            End If
                        End If
                        If CInt(r("is_descending_key")) <> 0 Then
                            sOut &= " desc"
                        End If
                        i += 1
                    End If
                End If
            Next
            If sOut <> "" Then
                sOut &= ")"
                If sInc <> "" Then
                    sOut &= " include (" & Mid(sInc, 2) & ")"
                End If
                If sWth <> "" Then
                    sOut &= sWth
                End If
                If sOn <> "" Then
                    sOut &= sOn
                End If
                sOut &= vbCrLf
            End If
            Return sOut
        End Get
    End Property

    Public ReadOnly Property IndexText(ByVal IndexName As String) As String
        Get
            Dim i As Integer = 0
            Dim j As Integer
            Dim s As String
            Dim sRest As String = ""
            Dim iClust As Integer = 0
            Dim sFG As String
            Dim sInc As String = ""
            Dim sWth As String = ""
            Dim sOut As String = ""
            Dim sOn As String = ""
            Dim qName As String = slib.QuoteIdentifier(IndexName)

            If dtIndexs Is Nothing Then
                Return ""
            End If

            If dtIndexs.Rows.Count = 0 Then
                Return ""
            End If

            For Each r As DataRow In dtIndexs.Rows
                If CInt(r("is_primary_key")) = 0 Then
                    If IndexName = slib.GetString(r("name")) Then
                        If i = 0 Then
                            iClust = slib.GetInteger(r("type"), 0)
                            sFG = slib.GetString(r("filegroup"))
                            sOut = "declare @o integer, @i integer, @t tinyint" & vbCrLf
                            sOut &= "       ,@c1 integer, @c2 integer" & vbCrLf
                            sOut &= "set @o = object_id('" & sTable & "')" & vbCrLf
                            sOut &= vbCrLf
                            sOut &= "select  @i = i.index_id" & vbCrLf
                            sOut &= "       ,@t = i.type" & vbCrLf
                            sOut &= "       ,@c1 = case when e.fill_factor is null then -1 else 0 end" & vbCrLf
                            sOut &= "from    sys.indexes i" & vbCrLf
                            sOut &= "join    sys.data_spaces d" & vbCrLf
                            sOut &= "on      d.data_space_id = i.data_space_id" & vbCrLf
                            sOut &= "left join" & vbCrLf
                            sOut &= "(" & vbCrLf
                            sOut &= "    select  '" & sFG & "' name" & vbCrLf

                            j = slib.GetInteger(r("FILL_FACTOR"), 0)
                            sOut &= "           ," & j & " fill_factor" & vbCrLf
                            s = slib.GetString(r("PAD_INDEX"))
                            sOut &= "           ," & slib.GetString(IIf(s = "NO", _
                                                 "0", "1")) & " is_padded" & vbCrLf
                            s = slib.GetString(r("IGNORE_DUP_KEY"))
                            sOut &= "           ," & slib.GetString(IIf(s = "NO", _
                                            "0", "1")) & " ignore_dup_key" & vbCrLf
                            s = slib.GetString(r("ALLOW_ROW_LOCKS"))
                            sOut &= "           ," & slib.GetString(IIf(s = "NO", _
                                           "0", "1")) & " allow_row_locks" & vbCrLf
                            s = slib.GetString(r("ALLOW_PAGE_LOCKS"))
                            sOut &= "           ," & slib.GetString(IIf(s = "NO", _
                                           "0", "1")) & " allow_page_locks" & vbCrLf
                            sOut &= ") e" & vbCrLf
                            sOut &= "on      e.name = d.name" & vbCrLf
                            sOut &= "and     e.fill_factor = i.fill_factor" & vbCrLf
                            sOut &= "and     e.is_padded = i.is_padded" & vbCrLf
                            sOut &= "and     e.ignore_dup_key = i.ignore_dup_key" & vbCrLf
                            sOut &= "and     e.allow_row_locks = i.allow_row_locks" & vbCrLf
                            sOut &= "and     e.allow_page_locks = i.allow_page_locks" & vbCrLf
                            sOut &= "where   i.object_id = @o" & vbCrLf
                            sOut &= "and     i.name = '" & IndexName & "'" & vbCrLf
                            sOut &= vbCrLf
                            sOut &= "if @c1 = -1" & vbCrLf
                            sOut &= "begin" & vbCrLf
                            sOut &= "    print 'changing index ''" & IndexName & "'''" & vbCrLf
                            sOut &= "    drop index " & qSchema & "." & qTable & "." & qName & vbCrLf
                            sOut &= "    set @i = null" & vbCrLf
                            sOut &= "end" & vbCrLf
                            sOut &= "else if @i is not null" & vbCrLf
                            sOut &= "begin" & vbCrLf
                            sOut &= "    select  @c1 = sum(1)" & vbCrLf
                            sOut &= "           ,@c2 = sum(case when x.keyorder is null then 0 else 1 end)" & vbCrLf
                            sOut &= "    from    sys.index_columns ic" & vbCrLf
                            sOut &= "    join    sys.columns c" & vbCrLf
                            sOut &= "    on      c.object_id = @o" & vbCrLf
                            sOut &= "    and     c.column_id = ic.column_id" & vbCrLf
                            sOut &= "    left join" & vbCrLf
                            sOut &= "    (" & vbCrLf
                            sOut &= "        select  " & slib.GetString(r("key_ordinal")) & " keyorder"
                            sOut &= ", '" & slib.GetString(r("ColumnName")) & "' ColumnName, "
                            If CInt(r("is_descending_key")) = 0 Then
                                sOut &= "0"
                            Else
                                sOut &= "1"
                            End If
                            sOut &= " Descending, "
                            If CInt(r("is_included_column")) = 0 Then
                                sOut &= "0"
                            Else
                                sOut &= "1"
                            End If
                            sOut &= " Included" & vbCrLf

                            sRest = "    ) x" & vbCrLf
                            sRest &= "    on      x.keyorder = ic.key_ordinal" & vbCrLf
                            sRest &= "    and     x.ColumnName = c.name" & vbCrLf
                            sRest &= "    and     x.Descending = ic.is_descending_key" & vbCrLf
                            sRest &= "    where   @t = " & slib.GetString(r("type")) & vbCrLf
                            sRest &= "    and     ic.object_id = @o" & vbCrLf
                            sRest &= "    and     ic.index_id = @i" & vbCrLf
                            sRest &= vbCrLf
                            sRest &= "    if @c1 <> @c2 or @c1 <> ~~" & vbCrLf
                            sRest &= "    begin" & vbCrLf
                            sRest &= "        print 'changing index ''" & IndexName & "'''" & vbCrLf
                            sRest &= "        drop index " & qSchema & "." & qTable & "." & qName & vbCrLf
                            sRest &= "        set @i = null" & vbCrLf
                            sRest &= "    end" & vbCrLf
                            sRest &= "end" & vbCrLf
                            sRest &= vbCrLf
                            sRest &= "if @i is null" & vbCrLf
                            sRest &= "begin" & vbCrLf
                            sRest &= "    print 'creating index ''" & IndexName & "'''" & vbCrLf
                            sRest &= "    create" & slib.GetString(IIf(CInt(r("is_unique")) <> 0, " unique", ""))
                            If iClust = 1 Then
                                sRest &= " clustered"
                            Else
                                If sFG <> sDefFileGroup Then
                                    sOn &= "on " & sFG & vbCrLf
                                End If
                                sRest &= " nonclustered"
                            End If
                            sRest &= " index " & qName & vbCrLf
                            sRest &= "      on " & qSchema & "." & qTable & " ("
                            sRest &= slib.QuoteIdentifier(r("ColumnName"))
                            If CInt(r("is_descending_key")) <> 0 Then
                                sRest &= " desc"
                            End If
                            sWth = IndexWith(r)
                        Else
                            sOut &= "        union select  " & slib.GetString(r("key_ordinal"))
                            sOut &= ", '" & slib.GetString(r("ColumnName")) & "', "
                            If CInt(r("is_descending_key")) = 0 Then
                                sOut &= "0"
                            Else
                                sOut &= "1"
                            End If
                            If CInt(r("is_included_column")) = 0 Then
                                sOut &= ", 0"
                            Else
                                sOut &= ", 1"
                            End If
                            sOut &= vbCrLf

                            If CInt(r("is_included_column")) = 0 Then
                                sRest &= "," & slib.QuoteIdentifier(r("ColumnName"))
                                If CInt(r("is_descending_key")) <> 0 Then
                                    sRest &= " desc"
                                End If
                            Else
                                sInc &= "," & slib.QuoteIdentifier(r("ColumnName"))
                                If CInt(r("is_descending_key")) <> 0 Then
                                    sInc &= " desc"
                                End If
                            End If
                        End If
                        i += 1
                    End If
                End If
            Next
            If sOut <> "" Then
                sOut &= sRest & ")" & vbCrLf
                If sInc <> "" Then
                    sOut &= "      include (" & Mid(sInc, 2) & ")" & vbCrLf
                End If
                If sWth <> "" Then
                    sOut &= "      " & sWth
                End If
                If sOn <> "" Then
                    sOut &= "      " & sOn
                End If
                sOut &= "end" & vbCrLf
                sOut = sOut.Replace("~~", Str(i))
            End If
            Return sOut
        End Get
    End Property

    Public ReadOnly Property PermissionText() As String
        Get
            Dim i As Integer
            Dim j As Integer
            Dim sOut As String = ""
            Dim s As String
            Dim sC As String

            If dtPerms Is Nothing Then
                Return ""
            End If

            If dtPerms.Rows.Count = 0 Then
                Return ""
            End If

            For Each r As DataRow In dtPerms.Rows
                s = LCase(slib.GetString(r.Item("permission_name")))
                j = slib.GetInteger(r.Item("columns"), 0)  ' insert/delete return null
                If j > 1 Then
                    sC = ""
                    s &= " ("
                    For i = 0 To Keys.GetUpperBound(0)
                        If (j And CInt(2 ^ (i + 1))) <> 0 Then
                            s &= sC & Keys(i)
                            sC = ", "
                        End If
                    Next
                    s &= ")"
                End If
                s &= " on " & qSchema & "." & qTable
                s &= " to " & slib.GetString(r.Item("grantee"))
                Select Case slib.GetString(r.Item("state"))
                    Case "GRANT_WITH_GRANT_OPTION"
                        sOut &= "grant " & s & " with grant option" & vbCrLf

                    Case "GRANT"
                        sOut &= "grant " & s & vbCrLf

                    Case "DENY"
                        sOut &= "deny " & s & vbCrLf

                End Select
            Next
            Return sOut
        End Get
    End Property

    Public ReadOnly Property LinkedTable(ByVal sFKeyName As String) As String
        Get
            Dim sFTable As String = ""

            If dtFKeys Is Nothing Then
                Return ""
            End If

            If dtFKeys.Rows.Count = 0 Then
                Return ""
            End If

            For Each r As DataRow In dtFKeys.Rows
                If sFKeyName = slib.GetString(r.Item("ConstraintName")) Then
                    sFTable = slib.GetString(r.Item("LinkedTable"))
                    Return sFTable
                End If
            Next
            Return sFTable
        End Get
    End Property

    Public ReadOnly Property FKeyText(ByVal sFKeyName As String) As String
        Get
            Dim i As Integer = 0
            Dim s As String
            Dim ss As String = ""
            Dim sOut As String = ""
            Dim sRest As String = ""
            Dim sOpt As String = ""
            Dim st As String = ""
            Dim sFTable As String = ""
            Dim qName As String = slib.QuoteIdentifier(sFKeyName)
            Dim bRepl As Boolean = True

            If dtFKeys Is Nothing Then
                Return ""
            End If

            If dtFKeys.Rows.Count = 0 Then
                Return ""
            End If

            For Each r As DataRow In dtFKeys.Rows
                If sFKeyName = slib.GetString(r.Item("ConstraintName")) Then
                    If i = 0 Then
                        bRepl = slib.GetBit(r("Replicated"), False)
                        sFTable = slib.GetString(r.Item("LinkedTable"))
                        sOut &= "declare @c1 integer, @c2 integer" & vbCrLf
                        sOut &= vbCrLf
                        sOut &= "if object_id('" & sFKeyName & "') is not null" & vbCrLf
                        sOut &= "begin" & vbCrLf
                        sOut &= "    select  @c1 = sum(1)" & vbCrLf
                        sOut &= "           ,@c2 = sum(case when x.keyno is null then 0 else 1 end)" & vbCrLf
                        sOut &= "    from    INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS c" & vbCrLf
                        sOut &= "    join    INFORMATION_SCHEMA.KEY_COLUMN_USAGE u1" & vbCrLf
                        sOut &= "    on      u1.CONSTRAINT_CATALOG = c.CONSTRAINT_CATALOG" & vbCrLf
                        sOut &= "    and     u1.CONSTRAINT_SCHEMA = c.CONSTRAINT_SCHEMA" & vbCrLf
                        sOut &= "    and     u1.CONSTRAINT_NAME = c.CONSTRAINT_NAME" & vbCrLf
                        sOut &= "    join    INFORMATION_SCHEMA.KEY_COLUMN_USAGE u2" & vbCrLf
                        sOut &= "    on      u2.CONSTRAINT_CATALOG = c.UNIQUE_CONSTRAINT_CATALOG" & vbCrLf
                        sOut &= "    and     u2.CONSTRAINT_SCHEMA = c.UNIQUE_CONSTRAINT_SCHEMA" & vbCrLf
                        sOut &= "    and     u2.CONSTRAINT_NAME = c.UNIQUE_CONSTRAINT_NAME" & vbCrLf
                        sOut &= "    join" & vbCrLf
                        sOut &= "    (" & vbCrLf
                        sOut &= "        select  " & CInt(r.Item("Sequence")) & " keyno, '"
                        sOut &= slib.GetString(r.Item("ColumnName")) & "' lkey, '"
                        sOut &= slib.GetString(r.Item("LinkedColumn")) & "' fkey" & vbCrLf

                        sRest = "    ) x" & vbCrLf
                        sRest &= "    on      x.keyno = u1.ORDINAL_POSITION" & vbCrLf
                        sRest &= "    and     x.lkey = u1.COLUMN_NAME" & vbCrLf
                        sRest &= "    and     x.fkey = u2.COLUMN_NAME" & vbCrLf
                        sRest &= "    where   c.CONSTRAINT_NAME = '" & sFKeyName & "'" & vbCrLf
                        s = slib.GetString(r.Item("DELETE_RULE"))
                        If s <> "NO ACTION" Then
                            sOpt = "on delete " & s
                            st = " "
                        End If
                        sRest &= "    and     c.DELETE_RULE = '" & s & "'" & vbCrLf
                        s = slib.GetString(r.Item("UPDATE_RULE"))
                        If s <> "NO ACTION" Then
                            sOpt &= st & "on update " & s
                        End If
                        sRest &= "    and     c.UPDATE_RULE = '" & s & "'" & vbCrLf
                        sRest &= "    and     objectproperty(object_id('" & qSchema & "." & qName & "'),'CnstIsNotRepl') = "
                        If bRepl Then
                            sRest &= "1" & vbCrLf
                        Else
                            sRest &= "0" & vbCrLf
                        End If
                        sRest &= "    and     u1.TABLE_NAME = '" & sTable & "'" & vbCrLf
                        sRest &= "    and     u2.TABLE_NAME = '" & sFTable & "'" & vbCrLf
                        sRest &= vbCrLf
                        sRest &= "    if coalesce(@c1,0) <> coalesce(@c2,0) or coalesce(@c1,0) <> ~~" & vbCrLf
                        sRest &= "    begin" & vbCrLf
                        sRest &= "        print 'changing foreign key ''" & sFKeyName & "'''" & vbCrLf
                        sRest &= "        alter table " & qSchema & "." & qTable & " drop constraint " & qName & vbCrLf
                        sRest &= "    end" & vbCrLf
                        sRest &= "end" & vbCrLf
                        sRest &= vbCrLf
                        sRest &= "if object_id('" & sFKeyName & "') is null" & vbCrLf
                        sRest &= "begin" & vbCrLf
                        sRest &= "    print 'creating foreign key ''" & sFKeyName & "'''" & vbCrLf
                        sRest &= "    alter table " & qSchema & "." & qTable & " add constraint " & qName & vbCrLf
                        sRest &= "    foreign key (" & slib.QuoteIdentifier(r.Item("ColumnName"))
                        ss = ") references " & slib.QuoteIdentifier(r.Item("LinkedSchema")) & "." & _
                             slib.QuoteIdentifier(r.Item("LinkedTable")) & "(" & slib.QuoteIdentifier(r.Item("LinkedColumn"))
                    Else
                        sOut &= "        union select  " & CInt(r.Item("Sequence")) & ", '" & slib.GetString(r.Item("ColumnName"))
                        sOut &= "', '" & slib.QuoteIdentifier(r.Item("LinkedColumn")) & "'" & vbCrLf
                        sRest &= "," & slib.QuoteIdentifier(r.Item("ColumnName"))
                        ss &= "," & slib.QuoteIdentifier(r.Item("LinkedColumn"))
                    End If
                    i += 1
                End If
            Next
            If sOut <> "" Then
                sOut &= sRest & ss & ")"
                If bRepl Then
                    sOut &= " not for replication"
                End If
                sOut &= vbCrLf
                If sOpt <> "" Then
                    sOut &= "    " & sOpt & vbCrLf
                End If
                sOut &= "end" & vbCrLf
                sOut = sOut.Replace("~~", Str(i))
            End If
            Return sOut
        End Get
    End Property

    Public ReadOnly Property FKeyShort(ByVal sFKeyName As String) As String
        Get
            Dim i As Integer = 0
            Dim s As String
            Dim ss As String = ""
            Dim so As String = ""
            Dim st As String = ""
            Dim sOut As String = ""
            Dim bRepl As Boolean = True

            If dtFKeys Is Nothing Then
                Return ""
            End If

            If dtFKeys.Rows.Count = 0 Then
                Return ""
            End If

            For Each r As DataRow In dtFKeys.Rows
                If sFKeyName = slib.GetString(r.Item("ConstraintName")) Then
                    If i = 0 Then
                        bRepl = slib.GetBit(r("Replicated"), False)
                        sOut &= "alter table " & qSchema & "." & qTable & " add constraint " & slib.QuoteIdentifier(sFKeyName) & vbCrLf
                        sOut &= "foreign key (" & slib.QuoteIdentifier(r.Item("ColumnName"))

                        ss = ") references " & slib.QuoteIdentifier(r.Item("LinkedSchema")) & "." & _
                            slib.QuoteIdentifier(r.Item("LinkedTable")) & "(" & slib.QuoteIdentifier(r.Item("LinkedColumn"))
                        s = slib.GetString(r.Item("DELETE_RULE"))
                        If s <> "NO ACTION" Then
                            so = "on delete " & s
                            st = " "
                        End If
                        s = slib.GetString(r.Item("UPDATE_RULE"))
                        If s <> "NO ACTION" Then
                            so &= st & "on update " & s
                        End If
                    Else
                        sOut &= "," & slib.QuoteIdentifier(r.Item("ColumnName"))
                        ss &= "," & slib.QuoteIdentifier(r.Item("LinkedColumn"))
                    End If
                    i += 1
                End If
            Next
            If sOut <> "" Then
                sOut &= ss & ")"
                If bRepl Then
                    sOut &= " not for replication"
                End If
                sOut &= vbCrLf
                If so <> "" Then
                    sOut &= so & vbCrLf
                End If
            End If
            Return sOut
        End Get
    End Property

    Public ReadOnly Property XML() As String
        Get
            Dim sOut As String
            Dim tc As TableColumn
            Dim i As Integer
            Dim s As String
            Dim ss As String
            Dim st As String
            Dim b As Boolean

            sOut = "<?xml version='1.0'?>" & vbCrLf
            sOut &= "<sqldef>" & vbCrLf
            sOut &= "  <table name='" & sTable & "' owner='" & sSchema & "'"
            If sFileGroup <> sDefFileGroup Then
                sOut &= " filegroup='" & sFileGroup & "'"
            End If
            If sTextFileGroup <> sFileGroup And sTextFileGroup <> "" Then
                sOut &= " textfilegroup='" & sTextFileGroup & "'"
            End If
            sOut &= ">" & vbCrLf

            sOut &= "    <columns>" & vbCrLf
            For Each s In Keys
                tc = DirectCast(Values.Item(s), TableColumn)
                ss = "      <column name='" & tc.Name & "'"
                If tc.Computed = "" Then
                    ss &= " type='" & tc.Type & "'"
                    If tc.Type = "xml" Then
                        If tc.XMLCollection <> "" Then
                            If tc.XMLDocument Then
                                ss &= " document='Y'"
                            Else
                                ss &= " content='Y'"
                            End If
                            ss &= " collection='" & tc.XMLCollection & "'"
                        End If
                    Else
                        If tc.Length > 0 Then
                            ss &= " length='" & tc.Length & "'"
                        ElseIf tc.Length = 0 Then
                            ss &= " length='max'"
                        End If
                        If tc.Precision > 0 Then
                            ss &= " precision='" & tc.Precision & "'"
                            ss &= " scale='" & tc.Scale & "'"
                        End If
                    End If
                    ss &= " allownulls='" & tc.Nullable & "'"
                    If tc.Identity Then
                        ss &= " seed='" & tc.Seed & "' increment='" & tc.Increment & "'"
                        If tc.Replicated Then
                            ss &= " replication='N'"
                        End If
                    End If
                    If tc.RowGuid Then
                        ss &= " rowguid='Y'"
                    End If
                    If tc.ANSIPadded = "N" Then
                        ss &= " ansipadded='N'"
                    End If
                    If bCollation And tc.Collation <> "" Then
                        If tc.Collation = sDefCollation Then
                            ss &= " collation='database_default'"
                        Else
                            ss &= " collation='" & tc.Collation & "'"
                        End If
                    End If
                    If tc.DefaultName <> "" Then
                        ss &= ">" & vbCrLf
                        ss &= "        <default "
                        If bConsName Then
                            ss &= "name='" & tc.DefaultName & "'"
                        End If
                        st = tc.DefaultValue
                        If Mid(st, 1, 1) = "(" And Right(st, 1) = ")" Then
                            st = Mid(st, 2, Len(st) - 2)
                        End If
                        ss &= "><![CDATA[" & st & "]]></default>" & vbCrLf
                        ss &= "      </column>"
                    Else
                        ss &= " />"
                    End If
                Else
                    ss &= " allownulls='" & tc.Nullable & "'"
                    If tc.ANSIPadded = "N" Then
                        ss &= " ansipadded='N'"
                    End If
                    If bCollation And tc.Collation <> "" Then
                        If tc.Collation = sDefCollation Then
                            ss &= " collation='database_default'"
                        Else
                            ss &= " collation='" & tc.Collation & "'"
                        End If
                    End If
                    If tc.Persisted Then
                        ss &= " persisted='Y'"
                    End If
                    ss &= ">" & vbCrLf
                    ss &= "        <formula><![CDATA[" & tc.Computed & "]]></formula>" & vbCrLf
                    ss &= "      </column>"
                End If
                sOut &= ss & vbCrLf
            Next
            sOut &= "    </columns>" & vbCrLf

            If sPKey <> "" Then
                ss = "    <primarykey"
                If bConsName Then
                    ss &= " name='" & sPKey & "'"
                End If
                ss &= " clustered='" & slib.GetString(IIf(bPKClust, "Y", "N")) & "'"

                If Not dtIndexs Is Nothing Then
                    For Each r As DataRow In dtIndexs.Rows
                        If CInt(r("is_primary_key")) <> 0 Then
                            i = slib.GetInteger(r.Item("FILL_FACTOR"), 0)
                            If i > 0 Then
                                ss &= " fillfactor='" & i & "'"
                            End If
                            st = slib.GetString(r.Item("PAD_INDEX"))
                            If st = "YES" Then
                                ss &= " pad='on'"
                            End If
                            st = slib.GetString(r.Item("IGNORE_DUP_KEY"))
                            If st = "YES" Then
                                ss &= " dup='on'"
                            End If
                            st = slib.GetString(r.Item("ALLOW_ROW_LOCKS"))
                            If st = "NO" Then
                                ss &= " rowlocks='off'"
                            End If
                            st = slib.GetString(r.Item("ALLOW_PAGE_LOCKS"))
                            If st = "NO" Then
                                ss &= " pagelocks='off'"
                            End If
                            If Not bPKClust Then
                                st = slib.GetString(r.Item("filegroup"))
                                If st <> sFileGroup Then
                                    ss &= " on='" & st & "'"
                                End If
                            End If
                            Exit For
                        End If
                    Next
                End If

                ss &= ">" & vbCrLf
                For Each s In xPKeys
                    tc = DirectCast(Values.Item(s), TableColumn)
                    ss &= "      <column name='" & tc.Name & "'"
                    If tc.Descend Then
                        ss &= " direction='desc'"
                    End If
                    ss &= " />" & vbCrLf
                Next
                ss &= "    </primarykey>" & vbCrLf
                sOut &= ss
            End If

            ss = ""
            For Each s In xIndexs
                If s <> "" And s <> sPKey Then
                    If Not dtIndexs Is Nothing Then
                        If ss = "" Then
                            ss = "    <indexes>" & vbCrLf
                        End If
                        ss &= "      <index name='" & s & "'"
                        b = True
                        For Each r As DataRow In dtIndexs.Rows
                            If s = slib.GetString(r.Item("name")) Then
                                If b Then
                                    b = False
                                    If CInt(r.Item("type")) = 1 Then
                                        ss &= " clustered='Y'"
                                    Else
                                        st = slib.GetString(r.Item("filegroup"))
                                        If st <> sDefFileGroup Then
                                            ss &= " on='" & st & "'"
                                        End If
                                    End If
                                    If CInt(r.Item("is_unique")) <> 0 Then
                                        ss &= " unique='Y'"
                                    End If
                                    i = slib.GetInteger(r.Item("FILL_FACTOR"), 0)
                                    If i > 0 Then
                                        ss &= " fillfactor='" & i & "'"
                                    End If
                                    st = slib.GetString(r.Item("PAD_INDEX"))
                                    If st = "YES" Then
                                        ss &= " pad='on'"
                                    End If
                                    st = slib.GetString(r.Item("IGNORE_DUP_KEY"))
                                    If st = "YES" Then
                                        ss &= " dup='on'"
                                    End If
                                    st = slib.GetString(r.Item("ALLOW_ROW_LOCKS"))
                                    If st = "NO" Then
                                        ss &= " rowlocks='off'"
                                    End If
                                    st = slib.GetString(r.Item("ALLOW_PAGE_LOCKS"))
                                    If st = "NO" Then
                                        ss &= " pagelocks='off'"
                                    End If
                                    ss &= ">" & vbCrLf
                                End If
                                ss &= "        <column name='" & slib.GetString(r.Item("ColumnName")) & "'"
                                If CInt(r.Item("is_included_column")) <> 0 Then
                                    ss &= " included='Y'"
                                End If
                                If CInt(r.Item("is_descending_key")) <> 0 Then
                                    ss &= " direction='desc'"
                                End If
                                ss &= " />" & vbCrLf
                            End If
                        Next
                        ss &= "      </index>" & vbCrLf
                    End If
                End If
            Next
            If ss <> "" Then
                ss &= "    </indexes>" & vbCrLf
                sOut &= ss
            End If

            ss = ""
            For Each dr As DataRow In dtCheck.Rows
                If ss = "" Then
                    ss = "    <constraints>" & vbCrLf
                End If
                ss &= "      <constraint "
                If bConsName Then
                    ss &= "name='" & slib.QuoteIdentifier(dr("CONSTRAINT_NAME")) & "' "
                End If
                If slib.GetBit(dr("is_not_for_replication"), False) Then
                    ss &= "replication='N' "
                End If
                ss &= "type='check'>" & vbCrLf
                ss &= "        <![CDATA["
                s = slib.GetString(dr("CHECK_CLAUSE"))
                If fixdef Then
                    s = FixCheckText(s)
                End If
                ss &= s
                ss &= "]]>" & vbCrLf
                ss &= "      </constraint>" & vbCrLf
            Next
            If ss <> "" Then
                ss &= "    </constraints>" & vbCrLf
                sOut &= ss
            End If

            ss = ""
            For Each s In xFKeys
                If s <> "" Then
                    If ss = "" Then
                        ss = "    <foreignkeys>" & vbCrLf
                    End If
                    ss &= "      <foreignkey name='" & s & "'"

                    b = True
                    For Each r As DataRow In dtFKeys.Rows
                        If s = slib.GetString(r.Item("ConstraintName")) Then
                            If b Then
                                b = False
                                ss &= " references='" & slib.GetString(r.Item("LinkedTable")) & "'"
                                st = slib.GetString(r.Item("DELETE_RULE"))
                                If st <> "NO ACTION" Then
                                    ss &= " ondelete='" & st & "'"
                                End If
                                st = slib.GetString(r.Item("UPDATE_RULE"))
                                If st <> "NO ACTION" Then
                                    ss &= " onupdate='" & st & "'"
                                End If
                                If slib.GetBit(r("Replicated"), False) Then
                                    ss &= " replication='N'"
                                End If
                                ss &= ">" & vbCrLf
                            End If
                            ss &= "        <column name='" & slib.GetString(r.Item("ColumnName")) & "'"
                            ss &= " linksto='" & slib.GetString(r.Item("LinkedColumn")) & "' />" & vbCrLf
                        End If
                    Next
                    ss &= "      </foreignkey>" & vbCrLf
                End If
            Next
            If ss <> "" Then
                ss &= "    </foreignkeys>" & vbCrLf
                sOut &= ss
            End If
            sOut &= "  </table>" & vbCrLf
            sOut &= "</sqldef>" & vbCrLf
            XML = sOut
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub New()
        PreLoad = 0
    End Sub

    Public Sub New(ByVal sTableName As String, ByVal sqllib As sql, ByVal bDef As Boolean)
        LoadTable(sTableName, "dbo", sqllib, bDef)
    End Sub

    Public Sub New(ByVal sTableName As String, ByVal Schema As String, _
                                ByVal sqllib As sql, ByVal bDef As Boolean)
        LoadTable(sTableName, Schema, sqllib, bDef)
    End Sub

    Private Sub LoadTable(ByVal sTableName As String, ByVal Sch As String, _
                                ByVal sqllib As sql, ByVal bDef As Boolean)
        Dim s As String = "a"
        Dim b As Boolean = False
        Dim sdn As String
        Dim sdv As String
        Dim sName As String
        Dim sPK As String
        Dim sType As String
        Dim sNull As String
        Dim sColl As String
        Dim sFormula As String
        Dim bPersist As Boolean
        Dim sAP As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim i As Integer
        Dim iSeed As Integer
        Dim iIncr As Integer
        Dim bRepl As Boolean = False

        sSchema = Sch
        qSchema = slib.QuoteIdentifier(Sch)
        slib = sqllib
        fixdef = bDef
        PreLoad = 2

        dt = slib.TableColumns(slib.QuoteIdentifier(sTableName), qSchema)
        If dt.Rows.Count = 0 Then
            PreLoad = 3
            Return
        End If

        sTable = sqllib.GetString(dt.Rows(0).Item("TableName"))
        qTable = sqllib.QuoteIdentifier(sTable)
        dr = slib.TableDetails(qTable, qSchema)
        If Not dr Is Nothing Then
            sIdentity = slib.GetString(dr("IdentityColumn"))
            iSeed = slib.GetInteger(dr("IdentitySeed"), 1)
            iIncr = slib.GetInteger(dr("IdentityIncrement"), 1)
            bRepl = slib.GetBit(dr("IdentityReplicated"), False)
            sFileGroup = slib.GetString(dr("DataFileGroup"))
            sTextFileGroup = slib.GetString(dr("TextFileGroup"))
            sDefFileGroup = slib.GetString(dr("DefFileGroup"))
            sDefCollation = slib.GetString(dr("DefCollation"))
            sPartitionScheme = slib.GetString(dr("PartitionScheme"))
            sSchemeColumn = slib.GetString(dr("SchemeColumn"))
        End If

        For Each dr In dt.Rows        ' Columns
            sName = sqllib.GetString(dr("COLUMN_NAME"))
            sType = sqllib.GetString(dr("DATA_TYPE"))
            sNull = Mid(sqllib.GetString(dr("IS_NULLABLE")), 1, 1)
            sdn = sqllib.GetString(dr("DEFAULT_NAME"))
            sdv = FixDefaultText(sqllib.GetString(dr("DEFAULT_TEXT")))
            sColl = sqllib.GetString(dr("COLLATION_NAME"))
            sAP = Mid(sqllib.GetString(dr("ANSIPadded")), 1, 1)
            If sName = sIdentity Then
                AddIdentityColumn(sName, sType, dr("CHARACTER_MAXIMUM_LENGTH"), _
                    dr.Item("NUMERIC_PRECISION"), dr("NUMERIC_SCALE"), sNull, _
                    iSeed, iIncr, sdn, sdv, sColl, sAP, bRepl)
            Else
                s = sqllib.GetString(dr("ROWGUID"))
                If s = "NO" Then
                    sAP = Mid(sqllib.GetString(dr("ANSIPadded")), 1, 1)
                    sFormula = FixCheckText(sqllib.GetString(dr("Computed")))
                    If sFormula = "" Then
                        If LCase(sType) = "xml" Then
                            AddXMLColumn(sName, dr("xmlschema"), _
                                dr("xmlcollection"), dr("is_xml_document"), sNull, _
                                sdn, sdv, sColl, sAP)
                        Else
                            AddColumn(sName, sType, dr("CHARACTER_MAXIMUM_LENGTH"), _
                                dr.Item("NUMERIC_PRECISION"), dr("NUMERIC_SCALE"), sNull, _
                                sdn, sdv, sColl, sAP)
                        End If
                    Else
                        If sqllib.GetString(dr("Persisted")) = "NO" Then
                            bPersist = False
                        Else
                            bPersist = True
                        End If
                        AddComputedColumn(sName, sFormula, sNull, bPersist)
                    End If
                Else
                    AddRowGuidColumn(sName, sType, dr("CHARACTER_MAXIMUM_LENGTH"), _
                        dr.Item("NUMERIC_PRECISION"), dr("NUMERIC_SCALE"), sNull, _
                        sdn, sdv, sAP)
                End If
            End If
        Next

        dtIndexs = slib.TableIndexes(qTable, qSchema)

        b = False
        sName = ""
        For Each dr In dtIndexs.Rows
            s = sqllib.GetString(dr("name"))
            i = CInt(dr.Item("type"))
            If CInt(dr.Item("is_primary_key")) <> 0 Then
                If Not b Then
                    sPKey = s
                    If i = 1 Then
                        bPKClust = True
                    Else
                        bPKClust = False
                    End If
                    b = True
                End If
                sPK = sqllib.GetString(dr("ColumnName"))
                If CInt(dr("is_descending_key")) <> 0 Then
                    AddPKey(sPK, True)
                Else
                    AddPKey(sPK, False)
                End If
            Else
                If sName <> s Then
                    sName = s
                    i = xIndexs.GetUpperBound(0)
                    If xIndexs(0) <> "" Then
                        i += 1
                        ReDim Preserve xIndexs(i)
                    End If
                    xIndexs(i) = sName
                End If
            End If
        Next

        dtCheck = slib.TableCheckConstraints(qTable, qSchema)

        dtFKeys = slib.TableFKeys(qTable, qSchema)
        sName = ""
        For Each r As DataRow In dtFKeys.Rows
            s = sqllib.GetString(r("ConstraintName"))
            If sName <> s Then
                sName = s
                i = xFKeys.GetUpperBound(0)
                If xFKeys(0) <> "" Then
                    i += 1
                    ReDim Preserve xFKeys(i)
                End If
                xFKeys(i) = sName
            End If
        Next

        dtPerms = slib.TablePermissions(qTable, qSchema)

        'dt = slib.TableTriggers(sTable)
        'For Each r As DataRow In dt.Rows
        '    i = xTriggers.GetUpperBound(0)
        '    If xTriggers(0) <> "" Then
        '        i += 1
        '        ReDim Preserve xTriggers(i)
        '    End If
        '    xTriggers(i) = sqllib.GetString(r("TriggerName"))
        'Next
    End Sub

    Public Sub AddColumn( _
        ByVal sName As String, _
        ByVal sType As String, _
        ByVal oLength As Object, _
        ByVal oPrecision As Object, _
        ByVal oScale As Object, _
        ByVal bNullable As String, _
        ByVal sDefaultName As String, _
        ByVal sDefaultValue As String, _
        ByVal sCollation As String, _
        ByVal sANSIPadded As String)

        Dim parm As New TableColumn

        With parm
            .Name = sName
            .Type = sType
            If IsNumeric(oLength) Then
                .Length = CInt(oLength)
            End If
            If IsNumeric(oPrecision) Then
                .Precision = CInt(oPrecision)
            End If
            If IsNumeric(oScale) Then
                .Scale = CInt(oScale)
            End If
            .Nullable = bNullable
            .DefaultName = sDefaultName
            .DefaultValue = sDefaultValue
            .Collation = sCollation
            .ANSIPadded = sANSIPadded
        End With

        AddColumn(parm)
    End Sub

    Public Sub AddXMLColumn( _
        ByVal sName As String, _
        ByVal oXMLSchema As Object, _
        ByVal oXMLCollection As Object, _
        ByVal oXMLDoc As Object, _
        ByVal bNullable As String, _
        ByVal sDefaultName As String, _
        ByVal sDefaultValue As String, _
        ByVal sCollation As String, _
        ByVal sANSIPadded As String)

        Dim parm As New TableColumn
        Dim s As String

        s = slib.QuoteIdentifier(oXMLSchema)
        If s <> "" Then s &= "."
        s &= slib.QuoteIdentifier(oXMLCollection)
        With parm
            .Name = sName
            .Type = "xml"
            .XMLCollection = s
            .XMLDocument = slib.GetBit(oXMLDoc, False)
            .Nullable = bNullable
            .DefaultName = sDefaultName
            .DefaultValue = sDefaultValue
            .Collation = sCollation
            .ANSIPadded = sANSIPadded
        End With

        AddColumn(parm)
    End Sub

    Public Sub AddComputedColumn( _
        ByVal sName As String, _
        ByVal sFormula As String, _
        ByVal bNullable As String, _
        ByVal bPersist As Boolean)

        Dim parm As New TableColumn

        With parm
            .Name = sName
            .Computed = sFormula
            .Persisted = bPersist
            .Nullable = bNullable
        End With

        AddColumn(parm)
    End Sub

    Public Sub AddRowGuidColumn( _
        ByVal sName As String, _
        ByVal sType As String, _
        ByVal oLength As Object, _
        ByVal oPrecision As Object, _
        ByVal oScale As Object, _
        ByVal bNullable As String, _
        ByVal sDefaultName As String, _
        ByVal sDefaultValue As String, _
        ByVal sANSIPadded As String)

        Dim parm As New TableColumn

        With parm
            .Name = sName
            .Type = sType
            If IsNumeric(oLength) Then
                .Length = CInt(oLength)
            End If
            If IsNumeric(oPrecision) Then
                .Precision = CInt(oPrecision)
            End If
            If IsNumeric(oScale) Then
                .Scale = CInt(oScale)
            End If
            .Nullable = bNullable
            .DefaultName = sDefaultName
            .DefaultValue = sDefaultValue
            .RowGuid = True
            .ANSIPadded = sANSIPadded
        End With

        AddColumn(parm)
    End Sub

    Public Sub AddIdentityColumn( _
        ByVal sName As String, _
        ByVal sType As String, _
        ByVal oLength As Object, _
        ByVal oPrecision As Object, _
        ByVal oScale As Object, _
        ByVal bNullable As String, _
        ByVal iSeed As Integer, _
        ByVal iIncr As Integer, _
        ByVal sDefaultName As String, _
        ByVal sDefaultValue As String, _
        ByVal sCollation As String, _
        ByVal sANSIPadded As String, _
        ByVal bRepl As Boolean)

        Dim parm As New TableColumn

        With parm
            .Name = sName
            .Type = sType
            If IsNumeric(oLength) Then
                .Length = CInt(oLength)
            End If
            If IsNumeric(oPrecision) Then
                .Precision = CInt(oPrecision)
            End If
            If IsNumeric(oScale) Then
                .Scale = CInt(oScale)
            End If
            .Nullable = bNullable
            .Identity = True
            .Seed = iSeed
            .Increment = iIncr
            .Replicated = bRepl
            .DefaultName = sDefaultName
            .DefaultValue = sDefaultValue
            .Collation = sCollation
            .ANSIPadded = sANSIPadded
        End With

        AddColumn(parm)
    End Sub

    Public Sub AddColumn(ByVal parm As TableColumn)
        If parm.Identity And sIdentity = "" Then
            sIdentity = parm.Name
        End If

        If parm.Name = "AuditID" Then
            bAudit = True
        End If
        If parm.Name = "State" Then
            bState = True
        End If

        With parm
            .Index = Values.Count
            .Primary = False
            .Descend = False
        End With

        If parm.Index > Keys.GetUpperBound(0) Then
            ReDim Preserve Keys(parm.Index)
        End If

        Values.Add(parm.Name, parm)
        Keys(parm.Index) = parm.Name
    End Sub

    Public Sub AddPKey(ByVal sKey As String, ByVal bDescend As Boolean)
        Dim i As Integer
        Dim tc As TableColumn

        i = xPKeys.GetUpperBound(0)
        If xPKeys(0) <> "" Then
            i += 1
            ReDim Preserve xPKeys(i)
        End If
        tc = DirectCast(Values.Item(sKey), TableColumn)
        tc.Primary = True
        tc.Descend = bDescend
        xPKeys(i) = sKey

        If PreLoad = 0 Then
            If sPKey = "" Then sPKey = sTable & "PK"
        End If
    End Sub

    Public Function DataScript(ByVal sFilter As String) As String
        Dim tc As TableColumn
        Dim dt As DataTable
        Dim sOut As String = ""
        Dim sHead As String
        Dim sTail As String = ""
        Dim s As String
        Dim cols As String
        Dim i As Integer
        Dim ss As String = ""

        If sIdentity <> "" Then
            sOut &= "set identity_insert " & qSchema & "." & qTable & " on" & vbCrLf
            sOut &= vbCrLf
        End If

        sHead = "insert into " & qSchema & "." & qTable & vbCrLf
        sHead &= "(" & vbCrLf
        s = "    "
        For Each cols In Keys
            tc = DirectCast(Values.Item(cols), TableColumn)
            sHead &= s & tc.QuotedName
            s = ", "
        Next
        sHead &= vbCrLf
        sHead &= ")" & vbCrLf
        s = "select  x."
        For Each cols In Keys
            tc = DirectCast(Values.Item(cols), TableColumn)
            sHead &= s & tc.QuotedName & vbCrLf
            s = "       ,x."
        Next
        sHead &= "from" & vbCrLf
        sHead &= "(" & vbCrLf

        i = 0
        dt = slib.TableData(sTable, sSchema, sFilter)
        For Each r As DataRow In dt.Rows
            If i = 0 Then
                sOut &= sTail
                sOut &= sHead
                s = "    select  "
                For Each cols In Keys
                    tc = DirectCast(Values.Item(cols), TableColumn)
                    sOut &= s & tc.DataFormat(r(tc.Name)) & " " & tc.QuotedName & vbCrLf
                    s = "           ,"
                Next
                sTail = ") x" & vbCrLf
                sTail &= "left join " & qSchema & "." & qTable & " a" & vbCrLf
                s = "on      a."
                For Each cols In xPKeys
                    tc = DirectCast(Values.Item(cols), TableColumn)
                    If ss = "" Then ss = tc.QuotedName
                    sTail &= s & tc.QuotedName & " = x." & tc.QuotedName & vbCrLf
                    s = "and     a."
                Next
                sTail &= "where   a." & ss & " is null" & vbCrLf
                sTail &= "go" & vbCrLf & vbCrLf

                i += 1
            Else
                s = "    union select "
                For Each cols In Keys
                    tc = DirectCast(Values.Item(cols), TableColumn)
                    sOut &= s & tc.DataFormat(r(tc.Name))
                    s = ", "
                Next
                sOut &= vbCrLf
                i += 1
            End If
            If i = 100 Then i = 0
        Next

        sOut &= sTail
        If sIdentity <> "" Then
            sOut &= vbCrLf
            sOut &= "set identity_insert " & qSchema & "." & qTable & " off" & vbCrLf
            sOut &= "go" & vbCrLf & vbCrLf
        End If

        Return sOut
    End Function
#End Region

#Region "private functions"
    Private Function CreateTable(ByVal bFull As Boolean) As String
        Dim sOut As String = ""
        Dim Comma As String
        Dim s As String
        Dim sTab As String
        Dim tc As TableColumn
        Dim bANSI As Boolean = False
        Dim bNonANSI As Boolean = False

        If bFull Then
            sTab = "    "
            sOut = "if object_id('" & qSchema & "." & qTable & "') is null" & vbCrLf
            sOut &= "begin" & vbCrLf
            sOut &= "    print 'creating " & sSchema & "." & sTable & "'" & vbCrLf
        Else
            sTab = ""
        End If

        If PreLoad = 3 Then
            Return ""
        End If

        For Each s In Keys
            tc = DirectCast(Values.Item(s), TableColumn)
            Select Case tc.ANSIPadded
                Case "Y"
                    bANSI = True
                Case "N"
                    bNonANSI = True
            End Select
        Next

        If bNonANSI Then  'not ANSI padding
            sOut &= vbCrLf
            If Not bANSI Then
                sOut &= sTab & "set ansi_padding off" & vbCrLf
            Else
                sOut &= "  -- columns exist with different ansi_padding settings" & vbCrLf
                sOut &= "  -- that have not been correctly scripted." & vbCrLf
            End If
            sOut &= vbCrLf
        End If

        sOut &= sTab & "create table " & qSchema & "." & qTable & vbCrLf
        sOut &= sTab & "(" & vbCrLf
        Comma = " "

        For Each s In Keys
            tc = DirectCast(Values.Item(s), TableColumn)
            sOut &= sTab & "   " & Comma & tc.QuotedName & " "
            If tc.Computed = "" Then
                sOut &= tc.TypeText
                If tc.Identity Then
                    sOut &= " identity(" & tc.Seed & "," & tc.Increment & ")"
                    If tc.Replicated Then
                        sOut &= " not for replication"
                    End If
                End If

                If tc.RowGuid Then
                    sOut &= " rowguidcol"
                End If

                If bCollation And tc.Collation <> "" Then
                    If tc.Collation = sDefCollation Then
                        sOut &= " collate database_default"
                    Else
                        sOut &= " collate " & tc.Collation
                    End If
                End If

                If tc.Nullable = "N" Then
                    sOut &= " not"
                End If
                sOut &= " null"
            Else
                sOut &= "as " & tc.Computed
                If tc.Persisted Then
                    sOut &= " persisted"
                    If tc.Nullable = "N" Then
                        sOut &= " not null"
                    End If
                End If
            End If

            If tc.DefaultName <> "" Then
                If bConsName Then
                    sOut &= " constraint " & tc.QuotedDefaultName
                End If
                sOut &= " default " & tc.DefaultValue
            End If

            If tc.ANSIPadded = "N" And bANSI Then
                sOut &= "   -- not ANSI"
            End If
            sOut &= vbCrLf
            Comma = ","
        Next

        If sPKey <> "" Then
            Comma = " "
            sOut &= sTab & "   ,"
            If bConsName Then
                sOut &= "constraint " & slib.QuoteIdentifier(sPKey) & " primary key"
            Else
                sOut &= "primary key"
            End If
            If bPKClust Then
                sOut &= " clustered"
            Else
                sOut &= " nonclustered"
            End If
            sOut &= vbCrLf
            sOut &= sTab & "    (" & vbCrLf
            For Each s In xPKeys
                tc = DirectCast(Values.Item(s), TableColumn)
                sOut &= sTab & "       " & Comma & tc.QuotedName
                If tc.Descend Then
                    sOut &= " desc"
                End If
                Comma = ","
                sOut &= vbCrLf
            Next
            sOut &= sTab & "    )"
            If Not dtIndexs Is Nothing Then
                For Each r As DataRow In dtIndexs.Rows
                    If CInt(r("is_primary_key")) <> 0 Then
                        s = IndexWith(r)
                        If s <> "" Then
                            sOut &= s
                        End If
                        If Not bPKClust Then
                            s = slib.GetString(r("filegroup"))
                            If s <> sFileGroup Then
                                sOut &= " on " & s & vbCrLf
                            End If
                        End If
                        Exit For
                    End If
                Next
            End If
            sOut &= vbCrLf
        End If

        If Not dtCheck Is Nothing Then
            For Each dr As DataRow In dtCheck.Rows
                sOut &= sTab & "   ,"
                If bConsName Then
                    sOut &= "constraint " & slib.QuoteIdentifier(dr("CONSTRAINT_NAME")) & " "
                End If
                s = slib.GetString(dr("CHECK_CLAUSE"))
                s = FixCheckText(s)
                sOut &= "check"
                If slib.GetBit(dr("is_not_for_replication"), False) Then
                    sOut &= " not for replication"
                End If
                sOut &= " (" & s & ")" & vbCrLf
            Next
        End If
        sOut &= sTab & ")"
        If sFileGroup <> sDefFileGroup Then
            sOut &= " on " & slib.QuoteIdentifier(sFileGroup)
        End If
        If sTextFileGroup <> sFileGroup And sTextFileGroup <> "" Then
            sOut &= " textimage_on " & slib.QuoteIdentifier(sTextFileGroup)
        End If
        sOut &= vbCrLf

        If bFull Then
            sOut &= "end" & vbCrLf
        End If
        Return sOut
    End Function

    Private Function IndexWith(ByRef r As DataRow) As String
        Dim i As Integer = 0
        Dim sWth As String = ""
        Dim s As String = ""
        Dim b As Boolean
        Dim sCm As String = ""

        i = slib.GetInteger(r("FILL_FACTOR"), 0)
        If i > 0 Then
            sWth = "fillfactor = " & i
            sCm = ", "
        End If
        b = slib.GetBit(r("no_recompute"), False)
        If b Then
            sWth &= sCm & "statistics_norecompute = on"
            sCm = ", "
        End If
        s = slib.GetString(r("PAD_INDEX"))
        If s = "YES" Then
            sWth &= sCm & "pad_index = on"
            sCm = ", "
        End If
        s = slib.GetString(r("IGNORE_DUP_KEY"))
        If s = "YES" Then
            sWth &= sCm & "ignore_dup_key = on"
            sCm = ", "
        End If
        s = slib.GetString(r("ALLOW_ROW_LOCKS"))
        If s = "NO" Then
            sWth &= sCm & "allow_row_locks = off"
            sCm = ", "
        End If
        s = slib.GetString(r("ALLOW_PAGE_LOCKS"))
        If s = "NO" Then
            sWth &= sCm & "allow_page_locks = off"
            sCm = ", "
        End If
        If sWth <> "" Then
            sWth = "with (" & sWth & ")" & vbCrLf
        End If
        Return sWth
    End Function

    Private Function FixDefaultText(ByVal sDefault As String) As String
        Dim s As String = ""
        Dim ss As String
        Dim sSave As String = ""
        Dim i As Integer
        Dim mode As Integer = 0

        For i = 1 To Len(sDefault)
            ss = Mid(sDefault, i, 1)
            Select Case mode
                Case 0
                    Select Case ss
                        Case "[", "]"

                        Case "'"
                            mode = 1
                            s &= ss

                        Case "("
                            If fixdef Then
                                If LCase(Right(s, 4)) <> "char" _
                                And LCase(Right(s, 7)) <> "decimal" _
                                And LCase(Right(s, 7)) <> "numeric" Then
                                    sSave = "("
                                    mode = 2
                                Else
                                    s &= ss
                                End If
                            Else
                                s &= ss
                            End If

                        Case Else
                            s &= ss

                    End Select

                Case 1
                    If ss = "'" Then mode = 0
                    s &= ss

                Case 2
                    Select Case ss
                        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "-", "+"
                            sSave &= ss

                        Case "("
                            s &= sSave
                            sSave = "("

                        Case ")"
                            If Len(sSave) > 1 Then
                                s &= Mid(sSave, 2)
                            Else
                                s &= "()"
                            End If
                            mode = 0

                        Case "[", "]"
                            s &= sSave
                            mode = 0

                        Case Else
                            s &= sSave & ss
                            mode = 0

                    End Select
            End Select
        Next
        If Mid(s, 1, 1) <> "(" Then s = "(" & s & ")"
        Return s
    End Function

    Private Function FixCheckText(ByVal sCheck As String) As String
        Dim s As String

        s = slib.RemoveSquares(sCheck)
        If Mid(s, 1, 1) = "(" And Right(s, 1) = ")" Then
            s = Mid(s, 2, Len(s) - 2)
        End If

        Return s
    End Function
#End Region
End Class

Option Explicit On
Option Strict On

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
    Private iLength As Integer = 0
    Private iPrecision As Integer = 0
    Private iScale As Integer = 0
    Private sqllib As New sql

#Region "Properties"
    Public Index As Integer
    Public Name As String
    Public Nullable As String
    Public Type As String
    Public DefaultName As String
    Public DefaultValue As String
    Public Primary As Boolean = False
    Public Descend As Boolean = False
    Public Identity As Boolean = False

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

    Public Property Length() As Integer
        Get
            Dim i As Integer
            Select Case Type
                Case "decimal", "numeric", "datetime", "smalldatetime"
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
            Select Case Type
                Case "char", "varchar", "nvarchar", "datetime", "smalldatetime", "sysname"
                    i = 0
                Case Else
                    i = iPrecision
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
            Select Case Type
                Case "char", "varchar", "nvarchar", "datetime", "smalldatetime", "sysname"
                    i = 0
                Case Else
                    i = iScale
            End Select
            Scale = i
        End Get
        Set(ByVal value As Integer)
            iScale = value
        End Set
    End Property

    Public ReadOnly Property vbType() As String
        Get
            Dim s As String
            Select Case Type
                Case "char", "varchar", "nvarchar", "sysname"
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
                    s = Type
            End Select
            vbType = s
        End Get
    End Property

    Public ReadOnly Property TypeText() As String
        Get
            Dim s As String
            s = Type
            Select Case Type
                Case "char", "varchar", "nvarchar"
                    s &= "(" & Length & ")"
                Case "decimal", "numeric"
                    s &= "(" & Precision & "," & Scale & ")"
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
    Private sPKey As String = ""
    Private bPKClust As Boolean = True
    Private sIdentity As String = ""
    Private bAudit As Boolean = False
    Private bState As Boolean = False
    Private bConsName As Boolean = True
    Private fixdef As Boolean = False

    Private xPKeys(0) As String
    Private xFKeys(0) As String
    Private xIndexs(0) As String
    Private dtFKeys As DataTable
    Private dtIndexs As DataTable

    Private Values As New Hashtable
    Private Keys(0) As String
    Private slib As New sql

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

    Public ReadOnly Property IndexText(ByVal IndexName As String) As String
        Get
            Dim i As Integer = 0
            Dim sRest As String = ""
            Dim sInc As String = ""
            Dim sOut As String = ""
            Dim qName As String = slib.QuoteIdentifier(IndexName)

            If dtIndexs Is Nothing Then
                Return ""
            End If

            If dtIndexs.Rows.Count = 0 Then
                Return ""
            End If

            For Each r As DataRow In dtIndexs.Rows
                If CInt(r.Item("is_primary_key")) = 0 Then
                    If IndexName = slib.GetString(r.Item("name")) Then
                        If i = 0 Then
                            sOut = "declare @o integer, @i integer, @t tinyint" & vbCrLf
                            sOut &= "       ,@c1 integer, @c2 integer" & vbCrLf
                            sOut &= "set @o = object_id('" & sTable & "')" & vbCrLf
                            sOut &= vbCrLf
                            sOut &= "select  @i = i.index_id" & vbCrLf
                            sOut &= "       ,@t = i.type" & vbCrLf
                            sOut &= "from    sys.indexes i" & vbCrLf
                            sOut &= "where   i.object_id = @o" & vbCrLf
                            sOut &= "and     i.name = '" & IndexName & "'" & vbCrLf
                            sOut &= vbCrLf
                            sOut &= "if @@rowcount > 0" & vbCrLf
                            sOut &= "begin" & vbCrLf
                            sOut &= "    select  @c1 = sum(1)" & vbCrLf
                            sOut &= "           ,@c2 = sum(case when x.keyorder is null then 0 else 1 end)" & vbCrLf
                            sOut &= "    from    sys.index_columns ic" & vbCrLf
                            sOut &= "    join    sys.columns c" & vbCrLf
                            sOut &= "    on      c.object_id = @o" & vbCrLf
                            sOut &= "    and     c.column_id = ic.column_id" & vbCrLf
                            sOut &= "    left join" & vbCrLf
                            sOut &= "    (" & vbCrLf
                            sOut &= "        select  " & slib.GetString(r.Item("key_ordinal")) & " keyorder"
                            sOut &= ", '" & slib.GetString(r.Item("ColumnName")) & "' ColumnName, "
                            If CInt(r.Item("is_descending_key")) = 0 Then
                                sOut &= "0"
                            Else
                                sOut &= "1"
                            End If
                            sOut &= " Descending, "
                            If CInt(r.Item("is_included_column")) = 0 Then
                                sOut &= "0"
                            Else
                                sOut &= "1"
                            End If
                            sOut &= " Included" & vbCrLf

                            sRest = "    ) x" & vbCrLf
                            sRest &= "    on      x.keyorder = ic.key_ordinal" & vbCrLf
                            sRest &= "    and     x.ColumnName = c.name" & vbCrLf
                            sRest &= "    and     x.Descending = ic.is_descending_key" & vbCrLf
                            sRest &= "    where   @t = " & slib.GetString(r.Item("type")) & vbCrLf
                            sRest &= "    and     ic.object_id = @o" & vbCrLf
                            sRest &= "    and     ic.index_id = @i" & vbCrLf
                            sRest &= vbCrLf
                            sRest &= "    if @c1 <> @c2 or @c1 <> ~~" & vbCrLf
                            sRest &= "    begin" & vbCrLf
                            sRest &= "        print 'changing index ''" & IndexName & "'''" & vbCrLf
                            sRest &= "        drop index dbo." & qTable & "." & qName & vbCrLf
                            sRest &= "        set @i = null" & vbCrLf
                            sRest &= "    end" & vbCrLf
                            sRest &= "end" & vbCrLf
                            sRest &= vbCrLf
                            sRest &= "if @i is null" & vbCrLf
                            sRest &= "begin" & vbCrLf
                            sRest &= "    print 'creating index ''" & IndexName & "'''" & vbCrLf
                            sRest &= "    create" & slib.GetString(IIf(CInt(r.Item("is_unique")) <> 0, " unique", ""))
                            sRest &= slib.GetString(IIf(CInt(r.Item("type")) = 1, " clustered", " nonclustered"))
                            sRest &= " index " & qName & vbCrLf
                            sRest &= "      on dbo." & qTable & " ("
                            sRest &= slib.QuoteIdentifier(r.Item("ColumnName"))
                        Else
                            sOut &= "        union select  " & slib.GetString(r.Item("key_ordinal"))
                            sOut &= ", '" & slib.GetString(r.Item("ColumnName")) & "', "
                            If CInt(r.Item("is_descending_key")) = 0 Then
                                sOut &= "0"
                            Else
                                sOut &= "1"
                            End If
                            If CInt(r.Item("is_included_column")) = 0 Then
                                sOut &= ", 0"
                            Else
                                sOut &= ", 1"
                            End If
                            sOut &= vbCrLf

                            If CInt(r.Item("is_included_column")) = 0 Then
                                sRest &= "," & slib.QuoteIdentifier(r.Item("ColumnName"))
                            Else
                                sInc &= "," & slib.QuoteIdentifier(r.Item("ColumnName"))
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
                sOut &= "end" & vbCrLf
                sOut = sOut.Replace("~~", Str(i))
            End If
            Return sOut
        End Get
    End Property

    Public ReadOnly Property IndexShort(ByVal IndexName As String) As String
        Get
            Dim i As Integer = 0
            Dim sOut As String = ""
            Dim sInc As String = ""

            If dtIndexs Is Nothing Then
                Return ""
            End If

            If dtIndexs.Rows.Count = 0 Then
                Return ""
            End If

            For Each r As DataRow In dtIndexs.Rows
                If CInt(r.Item("is_primary_key")) = 0 Then
                    If IndexName = slib.GetString(r.Item("name")) Then
                        If i = 0 Then
                            sOut &= "create" & slib.GetString(IIf(CInt(r.Item("is_unique")) <> 0, " unique", ""))
                            sOut &= slib.GetString(IIf(CInt(r.Item("type")) = 1, " clustered", " nonclustered"))
                            sOut &= " index " & slib.QuoteIdentifier(IndexName) & " on dbo." & qTable & " ("
                            sOut &= slib.QuoteIdentifier(r.Item("ColumnName"))
                        Else
                            If CInt(r.Item("is_included_column")) = 0 Then
                                sOut &= "," & slib.QuoteIdentifier(r.Item("ColumnName"))
                            Else
                                sInc &= "," & slib.QuoteIdentifier(r.Item("ColumnName"))
                            End If
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
                sOut &= vbCrLf
            End If
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
            Dim ss As String = ""
            Dim sOut As String = ""
            Dim sRest As String = ""
            Dim sFTable As String = ""
            Dim qName As String = slib.QuoteIdentifier(sFKeyName)

            If dtFKeys Is Nothing Then
                Return ""
            End If

            If dtFKeys.Rows.Count = 0 Then
                Return ""
            End If

            For Each r As DataRow In dtFKeys.Rows
                If sFKeyName = slib.GetString(r.Item("ConstraintName")) Then
                    If i = 0 Then
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
                        sOut &= "    and     u2.ORDINAL_POSITION = u1.ORDINAL_POSITION" & vbCrLf
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
                        sRest &= "    and     u1.TABLE_NAME = '" & sTable & "'" & vbCrLf
                        sRest &= "    and     u2.TABLE_NAME = '" & sFTable & "'" & vbCrLf
                        sRest &= vbCrLf
                        sRest &= "    if @c1 <> @c2 or @c1 <> ~~" & vbCrLf
                        sRest &= "    begin" & vbCrLf
                        sRest &= "        print 'changing foreign key ''" & sFKeyName & "'''" & vbCrLf
                        sRest &= "        alter table dbo." & qTable & " drop constraint " & qName & vbCrLf
                        sRest &= "    end" & vbCrLf
                        sRest &= "end" & vbCrLf
                        sRest &= vbCrLf
                        sRest &= "if object_id('" & sFKeyName & "') is null" & vbCrLf
                        sRest &= "begin" & vbCrLf
                        sRest &= "    print 'creating foreign key ''" & sFKeyName & "'''" & vbCrLf
                        sRest &= "    alter table dbo." & qTable & " add constraint " & qName & vbCrLf
                        sRest &= "    foreign key (" & slib.QuoteIdentifier(r.Item("ColumnName"))
                        ss = ") references dbo." & slib.QuoteIdentifier(r.Item("LinkedTable")) & "(" & slib.QuoteIdentifier(r.Item("LinkedColumn"))
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
                sOut &= sRest & ss & ")" & vbCrLf
                sOut &= "end" & vbCrLf
                sOut = sOut.Replace("~~", Str(i))
            End If
            Return sOut
        End Get
    End Property

    Public ReadOnly Property FKeyShort(ByVal sFKeyName As String) As String
        Get
            Dim i As Integer = 0
            Dim ss As String = ""
            Dim sOut As String = ""

            If dtFKeys Is Nothing Then
                Return ""
            End If

            If dtFKeys.Rows.Count = 0 Then
                Return ""
            End If

            For Each r As DataRow In dtFKeys.Rows
                If sFKeyName = slib.GetString(r.Item("ConstraintName")) Then
                    If i = 0 Then
                        sOut &= "alter table dbo." & qTable & " add constraint " & slib.QuoteIdentifier(sFKeyName) & vbCrLf
                        sOut &= "foreign key (" & slib.QuoteIdentifier(r.Item("ColumnName"))

                        ss = ") references dbo." & slib.QuoteIdentifier(r.Item("LinkedTable")) & "(" & slib.QuoteIdentifier(r.Item("LinkedColumn"))
                    Else
                        sOut &= "," & slib.QuoteIdentifier(r.Item("ColumnName"))
                        ss &= "," & slib.QuoteIdentifier(r.Item("LinkedColumn"))
                    End If
                    i += 1
                End If
            Next
            If sOut <> "" Then
                sOut &= ss & ")" & vbCrLf
            End If
            Return sOut
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub New()
        PreLoad = 0
    End Sub

    Public Sub New(ByVal sTableName As String, ByVal sqllib As sql, ByVal bDef As Boolean)
        Dim s As String = "a"
        Dim b As Boolean = False
        Dim sdn As String
        Dim sdv As String
        Dim sName As String
        Dim sPK As String
        Dim sType As String
        Dim sNull As String
        'Dim sIdentity As String = ""
        Dim dt As DataTable
        Dim dr As DataRow
        Dim i As Integer

        slib = sqllib
        fixdef = bDef
        PreLoad = 2

        dt = slib.TableColumns(sTableName)
        If dt.Rows.Count = 0 Then
            PreLoad = 3
            Return
        End If

        sTable = sqllib.GetString(dt.Rows(0).Item("TableName"))
        qTable = sqllib.QuoteIdentifier(sTable)
        sIdentity = slib.TableIdentity(sTable)

        For Each dr In dt.Rows        ' Columns
            sName = sqllib.GetString(dr.Item("COLUMN_NAME"))
            sType = sqllib.GetString(dr.Item("DATA_TYPE"))
            sNull = Mid(sqllib.GetString(dr.Item("IS_NULLABLE")), 1, 1)
            If sName = sIdentity Then
                b = True
            Else
                b = False
            End If
            sdn = sqllib.GetString(dr.Item("DEFAULT_NAME"))
            sdv = FixDefaultText(sqllib.GetString(dr.Item("DEFAULT_TEXT")))
            AddColumn(sName, sType, dr.Item("CHARACTER_MAXIMUM_LENGTH"), _
                dr.Item("NUMERIC_PRECISION"), dr.Item("NUMERIC_SCALE"), sNull, b, sdn, sdv)
        Next

        dtIndexs = slib.TableIndexes(sTable)

        b = False
        sName = ""
        For Each dr In dtIndexs.Rows
            s = sqllib.GetString(dr.Item("name"))
            If CInt(dr.Item("is_primary_key")) <> 0 Then
                If Not b Then
                    sPKey = s
                    i = CInt(dr.Item("type"))
                    If i = 1 Then
                        bPKClust = True
                    Else
                        bPKClust = False
                    End If
                    b = True
                End If
                sPK = sqllib.GetString(dr.Item("ColumnName"))
                If CInt(dr.Item("is_descending_key")) = 1 Then
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

        dtFKeys = slib.TableFKeys(sTable)
        sName = ""
        For Each r As DataRow In dtFKeys.Rows
            s = sqllib.GetString(r.Item("ConstraintName"))
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
        'Triggers = slib.Table(sTable)
    End Sub

    Public Sub AddColumn( _
        ByVal sName As String, _
        ByVal sType As String, _
        ByVal oLength As Object, _
        ByVal oPrecision As Object, _
        ByVal oScale As Object, _
        ByVal bNullable As String, _
        ByVal bIdentity As Boolean, _
        ByVal sDefaultName As String, _
        ByVal sDefaultValue As String)

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
            .Identity = bIdentity
            .DefaultName = sDefaultName
            .DefaultValue = sDefaultValue
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

        If parm.Type = "int" Then
            parm.Type = "integer"
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
            sOut &= "set identity_insert dbo." & sTable & " on" & vbCrLf
            sOut &= vbCrLf
        End If

        sHead = "insert into dbo." & sTable & vbCrLf
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
        dt = slib.TableData(sTable, sFilter)
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
                sTail &= "left join dbo." & sTable & " a" & vbCrLf
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
            sOut &= "set identity_insert dbo." & sTable & " off" & vbCrLf
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

        If bFull Then
            sTab = "    "
            sOut = "if object_id('dbo." & sTable & "') is null" & vbCrLf
            sOut &= "begin" & vbCrLf
            sOut &= "    print 'creating dbo." & sTable & "'" & vbCrLf
        Else
            sTab = ""
        End If

        sOut &= sTab & "create table dbo." & slib.QuoteIdentifier(sTable) & vbCrLf
        sOut &= sTab & "(" & vbCrLf
        Comma = " "

        For Each s In Keys
            tc = DirectCast(Values.Item(s), TableColumn)
            sOut &= sTab & "   " & Comma & tc.QuotedName & " " & tc.TypeText
            If tc.Identity Then
                sOut &= " identity(1, 1)"
            End If
            If tc.Nullable = "N" Then
                sOut &= " not"
            End If
            sOut &= " null"

            If tc.DefaultName <> "" Then
                If bConsName Then
                    sOut &= " constraint " & tc.QuotedDefaultName
                End If
                sOut &= " default " & tc.DefaultValue
            End If

            Comma = ","
            sOut &= vbCrLf
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
            sOut &= sTab & "    )" & vbCrLf
        End If

        sOut &= sTab & ")" & vbCrLf
        If bFull Then
            sOut &= "end" & vbCrLf
        End If
        Return sOut
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
#End Region
End Class

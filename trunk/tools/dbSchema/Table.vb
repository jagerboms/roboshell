Option Explicit On
Option Strict On

Imports System.Data.SqlClient

Public Class TableColumn
    Public Index As Integer
    Public Name As String
    Public Nullable As String
    Public Type As String
    Public Length As Integer = 0
    Public Precision As Integer = 0
    Public Scale As Integer = 0
    Public DefaultName As String
    Public DefaultValue As String
    Public TypeText As String
    Public Primary As Boolean = False
    Public Descend As Boolean = False
    Public Identity As Boolean = False
    Public vbType As String
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

    Public Sub New()
        PreLoad = 0
    End Sub

    Public Sub New(ByVal sTableName As String, ByVal sConnect As String, ByVal bDef As Boolean)
        Dim s As String = "a"
        Dim b As Boolean = False
        Dim sdn As String
        Dim sdv As String
        Dim sName As String
        Dim sPK As String
        Dim sType As String
        Dim sNull As String
        Dim sIdentity As String = ""
        Dim psConn As SqlConnection
        Dim dt As DataTable
        Dim dr As DataRow
        Dim i As Integer

        fixdef = bDef
        PreLoad = 2
        psConn = New SqlConnection(sConnect)
        AddHandler psConn.InfoMessage, AddressOf psConn_InfoMessage
        psConn.Open()

        dt = GetColumns(sTableName, psConn)
        If dt.Rows.Count = 0 Then
            PreLoad = 3
            Return
        End If

        sTable = GetString(dt.Rows(0).Item("TableName"))
        sIdentity = GetIdentityColumn(sTableName, psConn)

        For Each dr In dt.Rows        ' Columns
            sName = GetString(dr.Item("COLUMN_NAME"))
            sType = GetString(dr.Item("DATA_TYPE"))
            sNull = Mid(GetString(dr.Item("IS_NULLABLE")), 1, 1)
            If sName = sIdentity Then
                b = True
            Else
                b = False
            End If
            sdn = GetString(dr.Item("DEFAULT_NAME"))
            sdv = FixDefaultText(GetString(dr.Item("DEFAULT_TEXT")))
            AddColumn(sName, sType, dr.Item("CHARACTER_MAXIMUM_LENGTH"), _
                dr.Item("NUMERIC_PRECISION"), dr.Item("NUMERIC_SCALE"), sNull, b, sdn, sdv)
        Next

        dtIndexs = GetIndexes(sTable, psConn)

        b = False
        sName = ""
        For Each dr In dtIndexs.Rows
            s = GetString(dr.Item("name"))
            If CInt(dr.Item("is_primary_key")) = -1 Then
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
                sPK = GetString(dr.Item("ColumnName"))
                If CInt(dr.Item("is_descending_key")) = -1 Then
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

        dtFKeys = GetFKeys(sTable, psConn)
        sName = ""
        For Each r As DataRow In dtFKeys.Rows
            s = GetString(r.Item("ConstraintName"))
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
        'Triggers = TableDetails.Tables(5).Copy
        psConn.Close()
    End Sub

    Public Property TableName() As String
        Get
            TableName = sTable
        End Get
        Set(ByVal value As String)
            If PreLoad = 0 Then
                sTable = value
            End If
        End Set
    End Property

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
            parm.vbType = "integer"
        End If

        With parm
            .Index = Values.Count
            .TypeText = .Type
            .Primary = False
            .Descend = False
        End With

        Select Case parm.Type
            Case "char"
                parm.TypeText = parm.Type & "(" & parm.Length & ")"
                parm.vbType = "string"
                parm.Precision = 0
                parm.Scale = 0

            Case "varchar"
                parm.TypeText = parm.Type & "(" & parm.Length & ")"
                parm.vbType = "string"
                parm.Precision = 0
                parm.Scale = 0

            Case "nvarchar"
                parm.TypeText = parm.Type & "(" & parm.Length & ")"
                parm.vbType = "string"
                parm.Precision = 0
                parm.Scale = 0

            Case "decimal"
                parm.TypeText = parm.Type & "(" & parm.Precision & "," & parm.Scale & ")"
                parm.vbType = "double"
                parm.Length = 0

            Case "datetime"
                parm.vbType = "datetime"
                parm.Length = 0
                parm.Precision = 0
                parm.Scale = 0

            Case "smalldatetime"
                parm.vbType = "date"
                parm.Length = 0
                parm.Precision = 0
                parm.Scale = 0

            Case "sysname"
                parm.vbType = "string"
                parm.Length = 128
                parm.Precision = 0
                parm.Scale = 0

            Case Else
                parm.vbType = parm.Type
        End Select

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
        tc = CType(Values.Item(sKey), TableColumn)
        tc.Primary = True
        tc.Descend = bDescend
        xPKeys(i) = sKey

        If PreLoad = 0 Then
            If sPKey = "" Then sPKey = sTable & "PK"
        End If
    End Sub

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
                Return CType(Values.Item(index), TableColumn)
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
            Dim sOut As String = ""

            If dtIndexs Is Nothing Then
                Return ""
            End If

            If dtIndexs.Rows.Count = 0 Then
                Return ""
            End If

            For Each r As DataRow In dtIndexs.Rows
                If CInt(r.Item("is_primary_key")) = 0 Then
                    If IndexName = GetString(r.Item("name")) Then
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
                            sOut &= "        select  " & GetString(r.Item("index_column_id")) & " keyorder"
                            sOut &= ", '" & GetString(r.Item("ColumnName")) & "' ColumnName, "
                            If CInt(r.Item("is_descending_key")) = 0 Then
                                sOut &= "0"
                            Else
                                sOut &= "1"
                            End If
                            sOut &= " Descending" & vbCrLf

                            sRest = "    ) x" & vbCrLf
                            sRest &= "    on      x.keyorder = ic.index_column_id" & vbCrLf
                            sRest &= "    and     x.ColumnName = c.name" & vbCrLf
                            sRest &= "    and     x.Descending = ic.is_descending_key" & vbCrLf
                            sRest &= "    where   @t = " & GetString(r.Item("type")) & vbCrLf
                            sRest &= "    and     ic.object_id = @o" & vbCrLf
                            sRest &= "    and     ic.index_id = @i" & vbCrLf
                            sRest &= vbCrLf
                            sRest &= "    if @c1 <> @c2 or @c1 <> ~~" & vbCrLf
                            sRest &= "    begin" & vbCrLf
                            sRest &= "        print 'changing index ''" & IndexName & "'''" & vbCrLf
                            sRest &= "        drop index dbo." & sTable & "." & IndexName & vbCrLf
                            sRest &= "        set @i = null" & vbCrLf
                            sRest &= "    end" & vbCrLf
                            sRest &= "end" & vbCrLf
                            sRest &= vbCrLf
                            sRest &= "if @i is null" & vbCrLf
                            sRest &= "begin" & vbCrLf
                            sRest &= "    print 'creating index ''" & IndexName & "'''" & vbCrLf
                            sRest &= "    create" & GetString(IIf(CInt(r.Item("is_unique")) = -1, " unique", ""))
                            sRest &= GetString(IIf(CInt(r.Item("type")) = 1, " clustered", " nonclustered"))
                            sRest &= " index " & IndexName & " on dbo." & sTable & " ("
                            sRest &= GetString(r.Item("ColumnName"))
                        Else
                            sOut &= "        union select  " & GetString(r.Item("index_column_id"))
                            sOut &= ", '" & GetString(r.Item("ColumnName")) & "', "
                            If CInt(r.Item("is_descending_key")) = 0 Then
                                sOut &= "0"
                            Else
                                sOut &= "1"
                            End If
                            sOut &= vbCrLf

                            sRest &= "," & GetString(r.Item("ColumnName"))
                        End If
                        i += 1
                    End If
                End If
            Next
            If sOut <> "" Then
                sOut &= sRest & ")" & vbCrLf
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

            If dtIndexs Is Nothing Then
                Return ""
            End If

            If dtIndexs.Rows.Count = 0 Then
                Return ""
            End If

            For Each r As DataRow In dtIndexs.Rows
                If CInt(r.Item("is_primary_key")) = 0 Then
                    If IndexName = GetString(r.Item("name")) Then
                        If i = 0 Then
                            sOut &= "create" & GetString(IIf(CInt(r.Item("is_unique")) = -1, " unique", ""))
                            sOut &= GetString(IIf(CInt(r.Item("type")) = 1, " clustered", " nonclustered"))
                            sOut &= " index " & IndexName & " on dbo." & sTable & " ("
                            sOut &= GetString(r.Item("ColumnName"))
                        Else
                            sOut &= "," & GetString(r.Item("ColumnName"))
                        End If
                        i += 1
                    End If
                End If
            Next
            If sOut <> "" Then
                sOut &= ")" & vbCrLf
            End If
            Return sOut
        End Get
    End Property

    Public ReadOnly Property FKeyText(ByVal sFKeyName As String) As String
        Get
            Dim i As Integer = 0
            Dim ss As String = ""
            Dim sOut As String = ""
            Dim sRest As String = ""
            Dim sFTable As String = ""

            If dtFKeys Is Nothing Then
                Return ""
            End If

            If dtFKeys.Rows.Count = 0 Then
                Return ""
            End If

            For Each r As DataRow In dtFKeys.Rows
                If sFKeyName = GetString(r.Item("ConstraintName")) Then
                    If i = 0 Then
                        sFTable = GetString(r.Item("LinkedTable"))
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
                        sOut &= GetString(r.Item("ColumnName")) & "' lkey, '"
                        sOut &= GetString(r.Item("LinkedColumn")) & "' fkey" & vbCrLf

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
                        sRest &= "        alter table dbo." & sTable & " drop constraint " & sFKeyName & vbCrLf
                        sRest &= "    end" & vbCrLf
                        sRest &= "end" & vbCrLf
                        sRest &= vbCrLf
                        sRest &= "if object_id('" & sFKeyName & "') is null" & vbCrLf
                        sRest &= "begin" & vbCrLf
                        sRest &= "    print 'creating foreign key ''" & sFKeyName & "'''" & vbCrLf
                        sRest &= "    alter table dbo." & sTable & " add constraint " & sFKeyName & vbCrLf
                        sRest &= "    foreign key (" & GetString(r.Item("ColumnName"))
                        ss = ") references dbo." & GetString(r.Item("LinkedTable")) & "(" & GetString(r.Item("LinkedColumn"))
                    Else
                        sOut &= "        union select  " & CInt(r.Item("Sequence")) & ", '" & GetString(r.Item("ColumnName"))
                        sOut &= "', '" & GetString(r.Item("LinkedColumn")) & "'" & vbCrLf
                        sRest &= "," & GetString(r.Item("ColumnName"))
                        ss &= "," & GetString(r.Item("LinkedColumn"))
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
                If sFKeyName = GetString(r.Item("ConstraintName")) Then
                    If i = 0 Then
                        sOut &= "alter table dbo." & sTable & " add constraint " & sFKeyName & vbCrLf
                        sOut &= "foreign key (" & GetString(r.Item("ColumnName"))

                        ss = ") references dbo." & GetString(r.Item("LinkedTable")) & "(" & GetString(r.Item("LinkedColumn"))
                    Else
                        sOut &= "," & GetString(r.Item("ColumnName"))
                        ss &= "," & GetString(r.Item("LinkedColumn"))
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

        sOut &= sTab & "create table dbo." & sTable & vbCrLf
        sOut &= sTab & "(" & vbCrLf
        Comma = " "

        For Each s In Keys
            tc = CType(Values.Item(s), TableColumn)
            sOut &= sTab & "   " & Comma & FixFieldName(tc.Name) & " " & tc.TypeText
            If tc.Identity Then
                sOut &= " identity(1, 1)"
            End If
            If tc.Nullable = "N" Then
                sOut &= " not"
            End If
            sOut &= " null"

            If tc.DefaultName <> "" Then
                If bConsName Then
                    sOut &= " constraint " & tc.DefaultName
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
                sOut &= "constraint " & sPKey & " primary key"
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
                tc = CType(Values.Item(s), TableColumn)
                sOut &= sTab & "       " & Comma & FixFieldName(tc.Name)
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

#Region "common functions"
    Private Function GetString(ByVal objValue As Object) As String
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

    Private Sub psConn_InfoMessage(ByVal sender As Object, _
        ByVal e As System.Data.SqlClient.SqlInfoMessageEventArgs)
        Console.WriteLine(e.Message)
    End Sub

    Private Function GetColumns(ByVal sTableName As String, ByVal psConn As SqlConnection) As DataTable
        Dim s As String
        Dim psAdapt As SqlDataAdapter
        Dim TableDetails As New DataSet

        s = "declare @n sysname "
        s &= "set @n = '" & sTableName & "' "
        s &= "declare @o integer set @o = object_id(@n) "
        s &= "select object_name(@o) TableName"
        s &= ",i.COLUMN_NAME"
        s &= ",i.DATA_TYPE"
        s &= ",i.CHARACTER_MAXIMUM_LENGTH"
        s &= ",i.IS_NULLABLE"
        s &= ",i.NUMERIC_PRECISION"
        s &= ",i.NUMERIC_SCALE"
        s &= ",s.name DEFAULT_NAME"
        s &= ",i.COLUMN_DEFAULT DEFAULT_TEXT "
        s &= "from INFORMATION_SCHEMA.COLUMNS i "
        s &= "join dbo.syscolumns c "
        s &= "on c.id = @o "
        s &= "and c.name = i.COLUMN_NAME "
        s &= "left join dbo.sysobjects s "
        s &= "on s.id = c.cdefault "
        s &= "where TABLE_NAME = @n "
        s &= "and TABLE_SCHEMA = 'dbo' "
        s &= "order by ORDINAL_POSITION"

        psAdapt = New SqlDataAdapter(s, psConn)
        psAdapt.SelectCommand.CommandType = CommandType.Text
        psAdapt.Fill(TableDetails)

        Return TableDetails.Tables(0)
    End Function

    Private Function GetIndexes(ByVal sTableName As String, ByVal psConn As SqlConnection) As DataTable
        Dim s As String
        Dim psAdapt As SqlDataAdapter
        Dim TableDetails As New DataSet

        s = "select i.name,ic.index_column_id,c.name ColumnName,ic.is_descending_key,i.type,i.is_primary_key,i.is_unique "
        s &= "from sys.indexes i "
        s &= "join sys.index_columns ic "
        s &= "on ic.object_id = i.object_id "
        s &= "and ic.index_id = i.index_id "
        s &= "join sys.columns c "
        s &= "on c.object_id = i.object_id "
        s &= "and c.column_id = ic.column_id "
        s &= "where i.object_id = object_id('" & sTable & "') "
        s &= "order by 1, 2, 3"

        psAdapt = New SqlDataAdapter(s, psConn)
        psAdapt.SelectCommand.CommandType = CommandType.Text
        psAdapt.Fill(TableDetails)

        Return TableDetails.Tables(0)
    End Function

    Private Function GetIdentityColumn(ByVal sTableName As String, ByVal psConn As SqlConnection) As String
        Dim s As String
        Dim psAdapt As SqlDataAdapter
        Dim Details As New DataSet

        s = "select name ColName from syscolumns where id = object_id('" & sTableName & "') and colstat & 1 = 1"

        psAdapt = New SqlDataAdapter(s, psConn)
        psAdapt.SelectCommand.CommandType = CommandType.Text
        psAdapt.Fill(Details)

        s = ""
        If Details.Tables(0).Rows.Count > 0 Then
            s = GetString(Details.Tables(0).Rows(0).Item("ColName"))
        End If
        Return s
    End Function

    Private Function GetFKeys(ByVal sTableName As String, ByVal psConn As SqlConnection) As DataTable
        Dim s As String
        Dim psAdapt As SqlDataAdapter
        Dim TableDetails As New DataSet

        s = "select c.CONSTRAINT_NAME ConstraintName"
        s &= ",u1.ORDINAL_POSITION Sequence"
        s &= ",u1.COLUMN_NAME ColumnName"
        s &= ",u2.TABLE_NAME LinkedTable"
        s &= ",u2.COLUMN_NAME LinkedColumn "
        s &= "from INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS c "
        s &= "join INFORMATION_SCHEMA.KEY_COLUMN_USAGE u1 "
        s &= "on u1.CONSTRAINT_CATALOG = c.CONSTRAINT_CATALOG "
        s &= "and u1.CONSTRAINT_SCHEMA = c.CONSTRAINT_SCHEMA "
        s &= "and u1.CONSTRAINT_NAME = c.CONSTRAINT_NAME "
        s &= "join INFORMATION_SCHEMA.KEY_COLUMN_USAGE u2 "
        s &= "on u2.CONSTRAINT_CATALOG = c.UNIQUE_CONSTRAINT_CATALOG "
        s &= "and u2.CONSTRAINT_SCHEMA = c.UNIQUE_CONSTRAINT_SCHEMA "
        s &= "and u2.CONSTRAINT_NAME = c.UNIQUE_CONSTRAINT_NAME "
        s &= "and u2.ORDINAL_POSITION = u1.ORDINAL_POSITION "
        s &= "where u1.TABLE_NAME = '" & sTableName & "' "
        s &= "order by 1, 2"

        psAdapt = New SqlDataAdapter(s, psConn)
        psAdapt.SelectCommand.CommandType = CommandType.Text
        psAdapt.Fill(TableDetails)

        Return TableDetails.Tables(0)
    End Function

    Private Function FixFieldName(ByVal sField As String) As String
        Dim s As String = sField

        Select Case LCase(sField)
            Case "group", "percent", "key", "function", "deny", "order", "return"
                s = """" & sField & """"
            Case Else
                If InStr(sField, " ", CompareMethod.Text) > 0 Then
                    s = """" & sField & """"
                End If
        End Select

        Return s
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

        Return s
    End Function

    'Private Function GetTriggers(ByVal sTableName As String, ByVal psConn As SqlConnection) As DataTable
    '    Dim s As String
    '    Dim psAdapt As SqlDataAdapter
    '    Dim TableDetails As New DataSet

    '    s = "select o.name TriggerName "
    '    s &= "from dbo.sysobjects o "
    '    s &= "where o.type = 'TR' "
    '    s &= "and o.parent_obj = object_id('" & sTableName & "')"

    '    psAdapt = New SqlDataAdapter(s, psConn)
    '    psAdapt.SelectCommand.CommandType = CommandType.Text
    '    psAdapt.Fill(TableDetails)
    '    psConn.Close()

    '    Return TableDetails.Tables(0)
    'End Function
#End Region
End Class

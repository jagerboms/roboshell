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

    Public Sub New(ByVal sTableName As String, ByVal sConnect As String)
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
            sdv = GetString(dr.Item("DEFAULT_TEXT"))

            AddColumn(sName, sType, dr.Item("CHARACTER_MAXIMUM_LENGTH"), _
                dr.Item("NUMERIC_PRECISION"), dr.Item("NUMERIC_SCALE"), sNull, b, sdn, sdv)
        Next

        dtIndexs = GetIndexes(sTable, psConn)

        b = False
        sName = ""
        For Each dr In dtIndexs.Rows
            If GetString(dr.Item("PrimaryKey")) = "Y" Then
                If Not b Then
                    sPKey = GetString(dr.Item("IndexName"))
                    s = GetString(dr.Item("Cluster"))
                    If s <> "Y" Then
                        bPKClust = False
                    Else
                        bPKClust = True
                    End If
                    b = True
                End If
                sPK = GetString(dr.Item("ColumnName"))
                If GetString(dr.Item("Descending")) = "Y" Then
                    AddPKey(sPK, True)
                Else
                    AddPKey(sPK, False)
                End If
            Else
                s = GetString(dr.Item("IndexName"))
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
        If parm.DefaultName = "" Then
            parm.DefaultValue = ""
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
            Dim sOut As String
            Dim Comma As String
            Dim s As String
            Dim tc As TableColumn

            sOut = "if object_id('dbo." & sTable & "') is null" & vbCrLf
            sOut &= "begin" & vbCrLf
            sOut &= "    print 'creating dbo." & sTable & "'" & vbCrLf
            sOut &= "    create table dbo." & sTable & vbCrLf
            sOut &= "    (" & vbCrLf
            Comma = " "

            For Each s In Keys
                tc = CType(Values.Item(s), TableColumn)
                sOut &= "       " & Comma & tc.Name & " " & tc.TypeText
                If tc.Identity Then
                    sOut &= " identity(1, 1)"
                End If
                If tc.Nullable = "N" Then
                    sOut &= " not"
                End If
                sOut &= " null"

                If tc.DefaultName <> "" Then
                    sOut &= " constraint " & tc.DefaultName & " default " & tc.DefaultValue
                End If

                Comma = ","
                sOut &= vbCrLf
            Next

            If sPKey <> "" Then
                Comma = " "
                sOut &= "       ,constraint " & sPKey & " primary key"
                If bPKClust Then
                    sOut &= " clustered"
                End If
                sOut &= vbCrLf
                sOut &= "        (" & vbCrLf
                For Each s In xPKeys
                    tc = CType(Values.Item(s), TableColumn)
                    sOut &= "           " & Comma & tc.Name
                    If tc.Descend Then
                        sOut &= " desc"
                    End If
                    Comma = ","
                    sOut &= vbCrLf
                Next
                sOut &= "        )" & vbCrLf
            End If

            sOut &= "    )" & vbCrLf
            sOut &= "end" & vbCrLf

            Return sOut
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
                If GetString(r.Item("PrimaryKey")) <> "Y" Then
                    If IndexName = GetString(r.Item("IndexName")) Then
                        If i = 0 Then
                            sOut = "if (" & vbCrLf
                            sOut &= "    select  count(*)" & vbCrLf
                            sOut &= "    from    dbo.sysindexes i" & vbCrLf
                            sOut &= "    join    dbo.sysindexkeys k" & vbCrLf
                            sOut &= "    on      k.id = i.id" & vbCrLf
                            sOut &= "    and     k.indid = i.indid" & vbCrLf
                            sOut &= "    join" & vbCrLf
                            sOut &= "    (" & vbCrLf
                            sOut &= "        select  1 keyorder, '" & GetString(r.Item("ColumnName")) & "' ColumnName, '" & GetString(r.Item("Descending")) & "' Descending" & vbCrLf

                            sRest = "    ) x" & vbCrLf
                            sRest &= "    on      x.keyorder = k.keyno" & vbCrLf
                            sRest &= "    and     x.ColumnName = index_col(object_name(i.id), i.indid, k.keyno)" & vbCrLf
                            sRest &= "    and     x.Descending = case indexkey_property(i.id, i.indid, k.colid, 'isdescending') when 1 then 'Y' else 'N' end" & vbCrLf
                            sRest &= "    where   i.name = '" & IndexName & "'" & vbCrLf
                            sRest &= "    and     i.id = object_id('" & sTable & "')" & vbCrLf
                            sRest &= "    and" & vbCrLf
                            sRest &= "    (" & vbCrLf
                            sRest &= "        select  count(*)" & vbCrLf
                            sRest &= "        from    dbo.sysindexes ix" & vbCrLf
                            sRest &= "        where   ix.name = '" & IndexName & "'" & vbCrLf
                            sRest &= "        and     ix.id = object_id('" & sTable & "')" & vbCrLf
                            sRest &= "    ) = ~~" & vbCrLf
                            sRest &= ") <> ~~" & vbCrLf
                            sRest &= "begin" & vbCrLf
                            sRest &= "    if exists" & vbCrLf
                            sRest &= "    (" & vbCrLf
                            sRest &= "        select  'a'" & vbCrLf
                            sRest &= "        from    sysindexes o" & vbCrLf
                            sRest &= "        where   o.id = object_id('dbo." & sTable & "')" & vbCrLf
                            sRest &= "        and     o.name = '" & IndexName & "'" & vbCrLf
                            sRest &= "    )" & vbCrLf
                            sRest &= "    begin" & vbCrLf
                            sRest &= "        print 'changing index ''" & IndexName & "'''" & vbCrLf
                            sRest &= "        drop index dbo." & sTable & "." & IndexName & vbCrLf
                            sRest &= "    end" & vbCrLf
                            sRest &= "    else" & vbCrLf
                            sRest &= "    begin" & vbCrLf
                            sRest &= "        print 'creating index ''" & IndexName & "'''" & vbCrLf
                            sRest &= "    end" & vbCrLf
                            sRest &= "    create " & GetString(IIf(GetString(r.Item("UniqueIndex")) = "Y", "unique ", ""))
                            sRest &= GetString(IIf(GetString(r.Item("Cluster")) = "Y", "clustered", "nonclustered")) & " index " & IndexName & " on dbo." & sTable & " ("
                            sRest &= GetString(r.Item("ColumnName"))
                        Else
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
                If GetString(r.Item("PrimaryKey")) <> "Y" Then
                    If IndexName = GetString(r.Item("IndexName")) Then
                        If i = 0 Then
                            sOut &= "create " & GetString(IIf(GetString(r.Item("UniqueIndex")) = "Y", "unique ", ""))
                            sOut &= GetString(IIf(GetString(r.Item("Cluster")) = "Y", "clustered", "nonclustered")) & " index " & IndexName & " on dbo." & sTable & " ("
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

            If dtFKeys Is Nothing Then
                Return ""
            End If

            If dtFKeys.Rows.Count = 0 Then
                Return ""
            End If

            For Each r As DataRow In dtFKeys.Rows
                If sFKeyName = GetString(r.Item("ConstraintName")) Then
                    If i = 0 Then
                        sOut &= "if (" & vbCrLf
                        sOut &= "    select  count(*)" & vbCrLf
                        sOut &= "    from    dbo.sysforeignkeys k" & vbCrLf
                        sOut &= "    join" & vbCrLf
                        sOut &= "    (" & vbCrLf
                        sOut &= "        select  1 keyno, '" & GetString(r.Item("ColumnName")) & "' lkey, '" & GetString(r.Item("LinkedColumn")) & "' fkey" & vbCrLf

                        sRest = "    ) x" & vbCrLf
                        sRest &= "    on      x.keyno = k.keyno" & vbCrLf
                        sRest &= "    and     x.lkey = col_name(k.fkeyid, k.fkey)" & vbCrLf
                        sRest &= "    and     x.fkey = col_name(k.rkeyid, k.rkey)" & vbCrLf
                        sRest &= "    where   k.constid = object_id('" & sFKeyName & "')" & vbCrLf
                        sRest &= "    and     k.fkeyid = object_id('" & sTable & "')" & vbCrLf
                        sRest &= "    and     k.rkeyid = object_id('" & GetString(r.Item("LinkedTable")) & "')" & vbCrLf
                        sRest &= "    and" & vbCrLf
                        sRest &= "    (" & vbCrLf
                        sRest &= "        select  count(*)" & vbCrLf
                        sRest &= "        from    dbo.sysforeignkeys k" & vbCrLf
                        sRest &= "        where   k.constid = object_id('" & sFKeyName & "')" & vbCrLf
                        sRest &= "        and     k.fkeyid = object_id('" & sTable & "')" & vbCrLf
                        sRest &= "        and     k.rkeyid = object_id('" & GetString(r.Item("LinkedTable")) & "')" & vbCrLf
                        sRest &= "    ) = ~~" & vbCrLf
                        sRest &= ") <> ~~" & vbCrLf
                        sRest &= "begin" & vbCrLf
                        sRest &= "    if object_id('" & sFKeyName & "') is not null" & vbCrLf
                        sRest &= "    begin" & vbCrLf
                        sRest &= "        print 'changing foreign key ''" & sFKeyName & "'''" & vbCrLf
                        sRest &= "        alter table dbo." & sTable & " drop constraint " & sFKeyName & vbCrLf
                        sRest &= "    end" & vbCrLf
                        sRest &= "    else" & vbCrLf
                        sRest &= "    begin" & vbCrLf
                        sRest &= "        print 'creating foreign key ''" & sFKeyName & "'''" & vbCrLf
                        sRest &= "    end" & vbCrLf
                        sRest &= "    alter table dbo." & sTable & " add constraint " & sFKeyName & vbCrLf
                        sRest &= "    foreign key (" & GetString(r.Item("ColumnName"))

                        ss = ") references dbo." & GetString(r.Item("LinkedTable")) & "(" & GetString(r.Item("LinkedColumn"))
                    Else
                        sOut &= "        union select  " & CInt(r.Item("Sequence")) & ", '" & GetString(r.Item("ColumnName")) & "', '" & GetString(r.Item("LinkedColumn")) & "'" & vbCrLf
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
        s &= ",c.text DEFAULT_TEXT "
        s &= "from INFORMATION_SCHEMA.COLUMNS i "
        s &= "left join dbo.sysobjects s "
        s &= "on s.parent_obj = @o "
        s &= "and s.xtype = 'D' "
        s &= "and col_name(@o, s.info) = i.COLUMN_NAME "
        s &= "left join dbo.syscomments c "
        s &= "on c.id = s.id "
        s &= "and c.colid = 1 "
        s &= "where TABLE_NAME = @n "
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

        s = "select x.name IndexName"
        s &= ",i.keyno KeyOrder"
        s &= ",index_col(object_name(x.id), x.indid, i.keyno) ColumnName"
        s &= ",case indexkey_property(x.id, x.indid, i.colid, 'isdescending') when 1 then 'Y' else 'N' end Descending"
        s &= ",case when indexproperty(x.id, x.name, 'IsClustered') = 1 then 'Y' else 'N' end Cluster"
        s &= ",case when s.name is not null then 'Y' else 'N' end PrimaryKey"
        s &= ",case when indexproperty(x.id, x.name, 'IsUnique') = 1 then 'Y' else 'N' end UniqueIndex "
        s &= "from dbo.sysindexes x "
        s &= "join dbo.sysindexkeys i "
        s &= "on i.id = x.id "
        s &= "and i.indid = x.indid "
        s &= "left join dbo.sysobjects s "
        s &= "on s.name = x.name "
        s &= "and s.parent_obj = x.id "
        s &= "and s.xtype = 'PK' "
        s &= "where x.id = object_id('" & sTableName & "') "
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

        s = "declare @o integer set @o = object_id('" & sTableName & "') "
        s &= "select col_name(@o, column_id) ColName "
        s &= "from sys.identity_columns "
        s &= "where object_id = @o"

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

        s = "select object_name(k.constid) ConstraintName"
        s &= ",k.keyno Sequence"
        s &= ",col_name(k.fkeyid, k.fkey) ColumnName"
        s &= ",object_name(k.rkeyid) LinkedTable"
        s &= ",col_name(k.rkeyid, k.rkey) LinkedColumn "
        s &= "from dbo.sysforeignkeys k "
        s &= "where k.fkeyid = object_id('" & sTableName & "') "
        s &= "order by 1, 2"

        psAdapt = New SqlDataAdapter(s, psConn)
        psAdapt.SelectCommand.CommandType = CommandType.Text
        psAdapt.Fill(TableDetails)

        Return TableDetails.Tables(0)
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

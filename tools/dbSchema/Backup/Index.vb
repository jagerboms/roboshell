Option Explicit On
Option Strict On

' create [ unique ] [ clustered | nonclustered ] index INDEX_NAME
'     on <object> ( column [ ASC | DESC ] [ ,...n ] )
'     [ include ( column_name [ ,...n ] ) ]
'     [ with ( <relational_index_option> [ ,...n ] ) ]
'     [ on { filegroup_name
'          | partition_scheme_name ( column_name )
'          | default
'          }  ]
'
' <object> ::=
' {
'     [ database_name. [ schema_name ] . | schema_name. ]
'       table_or_view_name
' }
'
' <relational_index_option> ::=
' {
'     PAD_INDEX = { ON | OFF }
'   | FILLFACTOR = fillfactor
'   | SORT_IN_TEMPDB = { ON | OFF }
'   | IGNORE_DUP_KEY = { ON | OFF }
'   | STATISTICS_NORECOMPUTE = { ON | OFF }
'   | DROP_EXISTING = { ON | OFF }
'   | ONLINE = { ON | OFF }
'   | ALLOW_ROW_LOCKS = { ON | OFF }
'   | ALLOW_PAGE_LOCKS = { ON | OFF }
'   | MAXDOP = max_degree_of_parallelism
' }

Imports System
Imports System.Collections

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

Public Class IndexColumn
    Private sName As String
    Private bDescending As Boolean
    Private bIncluded As Boolean
    Private iKeyOrdinal As Integer
    Private iPartition As Integer

#Region "Properties"
    Public Property Name() As String
        Get
            Name = sName
        End Get
        Set(ByVal nm As String)
            sName = nm
        End Set
    End Property

    Public Property Ordinal() As Integer
        Get
            Ordinal = iKeyOrdinal
        End Get
        Set(ByVal ko As Integer)
            iKeyOrdinal = ko
        End Set
    End Property

    Public Property Descending() As Boolean
        Get
            Descending = bDescending
        End Get
        Set(ByVal bd As Boolean)
            bDescending = bd
        End Set
    End Property

    Public ReadOnly Property KeyColumn() As Boolean
        Get
            If iKeyOrdinal = 0 Then
                KeyColumn = False
            Else
                KeyColumn = True
            End If
        End Get
    End Property

    Public Property Included() As Boolean
        Get
            Included = bIncluded
        End Get
        Set(ByVal binc As Boolean)
            bIncluded = binc
        End Set
    End Property

    Public Property Partition() As Integer
        Get
            Partition = iPartition
        End Get
        Set(ByVal p As Integer)
            iPartition = p
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub New(ByVal ColumnName As String, _
                   ByVal KeyOrdinal As Integer, _
                   ByVal IsDescending As Boolean, _
                   ByVal IsIncluded As Boolean, _
                   ByVal PartitionOrdinal As Integer)
        sName = ColumnName
        iKeyOrdinal = KeyOrdinal
        bDescending = IsDescending
        bIncluded = IsIncluded
        iPartition = PartitionOrdinal
    End Sub
#End Region

End Class

Public Class IndexColumns
    Inherits CollectionBase

#Region "Properties"
    Default Public Overloads ReadOnly Property Item(ByVal Index As Integer) As IndexColumn
        Get
            Return CType(List.Item(Index), IndexColumn)
        End Get
    End Property

    Default Public Overloads ReadOnly Property Item(ByVal Name As String) As IndexColumn
        Get
            For Each ic As IndexColumn In Me
                If ic.Name = Name Then
                    Return ic
                End If
            Next
            Return Nothing
        End Get
    End Property
#End Region

#Region "Methods"
    Public Function Add(ByVal ColumnName As String, ByVal Ordinal As Integer, _
                        ByVal IsDescending As Boolean, ByVal IsIncluded As Boolean, _
                        ByVal PartOrdinal As Integer) As IndexColumn
        Dim ic As New IndexColumn(ColumnName, Ordinal, IsDescending, _
                         IsIncluded, PartOrdinal)
        List.Add(ic)
        Return ic
    End Function
#End Region

End Class

Public Class TableIndex
    Private sqllib As New sql
    Private sSchema As String
    Private qSchema As String
    Private sTable As String
    Private qTable As String
    Private sName As String = ""
    Private qName As String = ""
    Private bClustered As Boolean = False
    Private bPrimaryKey As Boolean = False
    Private bUnique As Boolean = False
    Private cFileGroup As New FileGroup
    Private iFillFactor As Integer = 0
    Private bPadIndex As Boolean = False
    Private bIgnoreDuplicates As Boolean = False
    Private bRowLocking As Boolean = True
    Private bPageLocking As Boolean = True
    Private bNoRecompute As Boolean = False
    Private bPKNamed As Boolean = False

    Private cCols As New IndexColumns

    ' partition scheme
    ' system named

#Region "Properties"
    Public Property Name() As String
        Get
            Name = sName
        End Get
        Set(ByVal value As String)
            sName = value
            qName = sqllib.QuoteIdentifier(sName)
        End Set
    End Property

    Public ReadOnly Property QuotedName() As String
        Get
            QuotedName = qName
        End Get
    End Property

    Public Property Clustered() As Boolean
        Get
            Clustered = bClustered
        End Get
        Set(ByVal value As Boolean)
            bClustered = value
        End Set
    End Property

    Public Property PrimaryKey() As Boolean
        Get
            PrimaryKey = bPrimaryKey
        End Get
        Set(ByVal value As Boolean)
            bPrimaryKey = value
        End Set
    End Property

    Public Property Unique() As Boolean
        Get
            Unique = bUnique
        End Get
        Set(ByVal value As Boolean)
            bUnique = value
        End Set
    End Property

    Public Property IndexFileGroup() As FileGroup
        Get
            IndexFileGroup = cFileGroup
        End Get
        Set(ByVal fg As FileGroup)
            cFileGroup = fg
        End Set
    End Property

    Public Property FillFactor() As Integer
        Get
            FillFactor = iFillFactor
        End Get
        Set(ByVal value As Integer)
            iFillFactor = value
        End Set
    End Property

    Public Property PadIndex() As Boolean
        Get
            PadIndex = bPadIndex
        End Get
        Set(ByVal value As Boolean)
            bPadIndex = value
        End Set
    End Property

    Public Property IgnoreDuplicates() As Boolean
        Get
            IgnoreDuplicates = bIgnoreDuplicates
        End Get
        Set(ByVal value As Boolean)
            bIgnoreDuplicates = value
        End Set
    End Property

    Public Property RowLocking() As Boolean
        Get
            RowLocking = bRowLocking
        End Get
        Set(ByVal value As Boolean)
            bRowLocking = value
        End Set
    End Property

    Public Property PageLocking() As Boolean
        Get
            PageLocking = bPageLocking
        End Get
        Set(ByVal value As Boolean)
            bPageLocking = value
        End Set
    End Property

    Public Property NoRecompute() As Boolean
        Get
            NoRecompute = bNoRecompute
        End Get
        Set(ByVal value As Boolean)
            bNoRecompute = value
        End Set
    End Property

    Public Property PrimaryKeyNamed() As Boolean
        Get
            PrimaryKeyNamed = bPKNamed
        End Get
        Set(ByVal pkn As Boolean)
            bPKNamed = pkn
        End Set
    End Property

    Public Property Columns() As IndexColumns
        Get
            Columns = cCols
        End Get
        Set(ByVal value As IndexColumns)
            cCols = value
        End Set
    End Property

    Public ReadOnly Property IndexShort() As String
        Get
            Dim s As String
            Dim sOut As String

            If bPrimaryKey Then
                Return ""     ' no script for primary key as it is part of the table definition
            End If

            sOut = "create"
            If bUnique Then
                sOut &= " unique"
            End If
            If bClustered Then
                sOut &= " clustered"
            Else
                sOut &= " nonclustered"
            End If
            sOut &= " index " & qName & " on " & qSchema & "." & qTable & " ("
            s = ""
            For Each r As IndexColumn In cCols
                If r.KeyColumn Then
                    sOut &= s & sqllib.QuoteIdentifier(r.Name)
                    If r.Descending Then
                        sOut &= " desc"
                    End If
                    s = ","
                End If
            Next
            sOut &= ")"

            s = ""
            For Each r As IndexColumn In cCols
                If r.Included Then
                    If s = "" Then
                        sOut &= " include ("
                    End If
                    sOut &= s & sqllib.QuoteIdentifier(r.Name)
                    s = ","
                End If
            Next
            If s <> "" Then
                sOut &= ")"
            End If

            s = IndexWith()
            If s <> "" Then
                sOut &= vbCrLf & s
            End If

            If Not bClustered Then
                sOut &= cFileGroup.IndexText
            End If

            sOut &= vbCrLf

            IndexShort = sOut
        End Get
    End Property

    Public ReadOnly Property IndexText() As String
        Get
            Dim s As String
            Dim sOut As String
            Dim i As Integer

            If bPrimaryKey Then
                Return ""     ' no script for primary key as it is part of the table definition
            End If

            sOut = "declare @o integer, @i integer, @t tinyint" & vbCrLf
            sOut &= "       ,@c1 integer, @c2 integer" & vbCrLf
            sOut &= "set @o = object_id('" & sSchema & "." & sTable & "')" & vbCrLf  'rha add schema
            sOut &= vbCrLf
            sOut &= "select  @i = i.index_id" & vbCrLf
            sOut &= "       ,@t = i.type" & vbCrLf
            sOut &= "       ,@c1 = case when e.fill_factor is null then -1 else 0 end" & vbCrLf
            sOut &= "from    sys.indexes i" & vbCrLf
            sOut &= "join    sys.data_spaces d" & vbCrLf
            sOut &= "on      d.data_space_id = i.data_space_id" & vbCrLf
            sOut &= "left join" & vbCrLf
            sOut &= "(" & vbCrLf
            sOut &= "    select  '" & cFileGroup.IndexGroup & "' name" & vbCrLf
            sOut &= "           ," & iFillFactor & " fill_factor" & vbCrLf
            sOut &= "           ," & Bool2Int(bPadIndex) & " is_padded" & vbCrLf
            sOut &= "           ," & Bool2Int(bIgnoreDuplicates) & " ignore_dup_key" & vbCrLf
            sOut &= "           ," & Bool2Int(bRowLocking) & " allow_row_locks" & vbCrLf
            sOut &= "           ," & Bool2Int(bPageLocking) & " allow_page_locks" & vbCrLf
            sOut &= ") e" & vbCrLf
            sOut &= "on      e.name = d.name" & vbCrLf
            sOut &= "and     e.fill_factor = i.fill_factor" & vbCrLf
            sOut &= "and     e.is_padded = i.is_padded" & vbCrLf
            sOut &= "and     e.ignore_dup_key = i.ignore_dup_key" & vbCrLf
            sOut &= "and     e.allow_row_locks = i.allow_row_locks" & vbCrLf
            sOut &= "and     e.allow_page_locks = i.allow_page_locks" & vbCrLf
            sOut &= "where   i.object_id = @o" & vbCrLf
            sOut &= "and     i.name = '" & sName & "'" & vbCrLf
            sOut &= vbCrLf
            sOut &= "if @c1 = -1" & vbCrLf
            sOut &= "begin" & vbCrLf
            sOut &= "    print 'changing index ''" & sName & "'''" & vbCrLf
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

            s = ""
            For Each r As IndexColumn In cCols
                If s = "" Then
                    sOut &= "        select  1 keyorder"
                    sOut &= ", '" & r.Name & "' ColumnName, "
                    sOut &= Bool2Int(r.Descending) & " Descending, "
                    sOut &= Bool2Int(r.Included) & " Included" & vbCrLf
                    s = ","
                    i = 1
                Else
                    sOut &= "        union select  " & i
                    sOut &= ", '" & r.Name
                    sOut &= "', " & Bool2Int(r.Descending)
                    sOut &= "', " & Bool2Int(r.Included) & vbCrLf
                End If
                i += 1
            Next
            sOut &= "    ) x" & vbCrLf
            sOut &= "    on      x.keyorder = ic.key_ordinal" & vbCrLf
            sOut &= "    and     x.ColumnName = c.name" & vbCrLf
            sOut &= "    and     x.Descending = ic.is_descending_key" & vbCrLf
            sOut &= "    where   @t = "
            If bClustered Then
                sOut &= "1"
            Else
                sOut &= "2"
            End If
            sOut &= vbCrLf
            sOut &= "    and     ic.object_id = @o" & vbCrLf
            sOut &= "    and     ic.index_id = @i" & vbCrLf
            sOut &= vbCrLf
            sOut &= "    if @c1 <> @c2 or @c1 <> " & cCols.Count & vbCrLf
            sOut &= "    begin" & vbCrLf
            sOut &= "        print 'changing index ''" & sName & "'''" & vbCrLf
            sOut &= "        drop index " & qSchema & "." & qTable & "." & qName & vbCrLf
            sOut &= "        set @i = null" & vbCrLf
            sOut &= "    end" & vbCrLf
            sOut &= "end" & vbCrLf
            sOut &= vbCrLf
            sOut &= "if @i is null" & vbCrLf
            sOut &= "begin" & vbCrLf
            sOut &= "    print 'creating index ''" & sName & "'''" & vbCrLf
            sOut &= "    create"
            If bUnique Then
                sOut &= " unique"
            End If
            If bClustered Then
                sOut &= " clustered"
            Else
                sOut &= " nonclustered"
            End If
            sOut &= " index " & qName & vbCrLf
            sOut &= "      on " & qSchema & "." & qTable & " ("
            s = ""
            For Each r As IndexColumn In cCols
                If r.KeyColumn Then
                    sOut &= s & sqllib.QuoteIdentifier(r.Name)
                    If r.Descending Then
                        sOut &= " desc"
                    End If
                    s = ","
                End If
            Next
            sOut &= ")"

            s = ""
            For Each r As IndexColumn In cCols
                If r.Included Then
                    If s = "" Then
                        sOut &= " include ("
                    End If
                    sOut &= s & sqllib.QuoteIdentifier(r.Name)
                    s = ","
                End If
            Next
            If s <> "" Then
                sOut &= ")"
            End If
            sOut &= vbCrLf

            s = IndexWith()
            If s <> "" Then
                sOut &= "      " & s & vbCrLf
            End If

            If Not bClustered Then
                s = cFileGroup.IndexText
                If s <> "" Then
                    sOut &= "     " & s & vbCrLf
                End If
            End If

            sOut &= "end" & vbCrLf

            IndexText = sOut
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub New(ByVal Schema As String, ByVal Table As String, ByVal IndexName As String)
        Me.Name = IndexName
        sSchema = Schema
        qSchema = sqllib.QuoteIdentifier(sSchema)
        sTable = Table
        qTable = sqllib.QuoteIdentifier(sTable)
    End Sub

    Public Function XMLText(ByVal sTab As String) As String
        Dim sOut As String

        If bPrimaryKey Then
            Return ""
        End If

        sOut = sTab & "<index name='" & sName & "'"
        If bClustered Then
            sOut &= " clustered='Y'"
        Else
            sOut &= cFileGroup.IndexXML
        End If
        If bUnique Then
            sOut &= " unique='Y'"
        End If
        If iFillFactor > 0 Then
            sOut &= " fillfactor='" & iFillFactor & "'"
        End If
        If bNoRecompute Then
            sOut &= " norecompute='Y'"
        End If
        If bPadIndex Then
            sOut &= " pad='on'"
        End If
        If bIgnoreDuplicates Then
            sOut &= " dup='on'"
        End If
        If Not bRowLocking Then
            sOut &= " rowlocks='off'"
        End If
        If Not bPageLocking Then
            sOut &= " pagelocks='off'"
        End If
        sOut &= ">" & vbCrLf
        For Each r As IndexColumn In cCols
            If r.KeyColumn Or r.Included Then
                sOut &= sTab & "  <column name='" & r.Name & "'"
                If r.Included Then
                    sOut &= " included='Y'"
                Else
                    If r.Descending Then
                        sOut &= " direction='desc'"
                    End If
                End If
                sOut &= " />" & vbCrLf
            End If
        Next
        sOut &= sTab & "</index>" & vbCrLf
        Return sOut
    End Function

    Public Function PrimaryKeyText(ByVal sTab As String, ByVal opt As ScriptOptions) As String
        Dim s As String
        Dim sC As String
        Dim sOut As String

        If Not bPrimaryKey Then
            Return ""
        End If

        sOut = sTab & "   ,"
        If opt.PrimaryKeyShowName And Not bPKNamed Then
            sOut &= "constraint " & qName & " "
        End If
        sOut &= "primary key"
        If opt.TargetEnvironment <> ScriptOptions.TargetEnvironments.PostGres Then
            If bClustered Then
                sOut &= " clustered"
            Else
                sOut &= " nonclustered"
            End If
        End If
        sOut &= vbCrLf

        sC = " "
        sOut &= sTab & "    (" & vbCrLf
        For Each r As IndexColumn In cCols
            sOut &= sTab & "       " & sC & sqllib.QuoteIdentifier(r.Name)
            If r.Descending Then
                sOut &= " desc"
            End If
            sOut &= vbCrLf
            sC = ","
        Next
        sOut &= sTab & "    )"

        s = IndexWith()
        If s <> "" Then
            sOut &= s
        End If

        If Not bClustered Then
            sOut &= cFileGroup.IndexText
        End If
        sOut &= vbCrLf
        Return sOut
    End Function

    Public Function PrimaryKeyXML(ByVal sTab As String, ByVal opt As ScriptOptions) As String
        Dim sOut As String

        If Not bPrimaryKey Then
            Return ""
        End If

        sOut = sTab & "<primarykey"
        If opt.PrimaryKeyShowName And Not bPKNamed Then
            sOut &= " name='" & sName & "'"
        End If
        sOut &= " clustered='"
        If bClustered Then
            sOut &= "Y'"
        Else
            sOut &= "N'"
        End If
        If iFillFactor > 0 Then
            sOut &= " fillfactor='" & iFillFactor & "'"
        End If
        If bPadIndex Then
            sOut &= " pad='on'"
        End If
        If bIgnoreDuplicates Then
            sOut &= " dup='on'"
        End If
        If Not bRowLocking Then
            sOut &= " rowlocks='off'"
        End If
        If Not bPageLocking Then
            sOut &= " pagelocks='off'"
        End If
        If Not bClustered Then
            sOut &= cFileGroup.IndexXML
        End If
        sOut &= ">" & vbCrLf

        For Each r As IndexColumn In cCols
            sOut &= sTab & "  <column name='" & r.Name & "'"
            If r.Descending Then
                sOut &= " direction='desc'"
            End If
            sOut &= " />" & vbCrLf
        Next
        sOut &= sTab & "</primarykey>" & vbCrLf
        Return sOut
    End Function
#End Region

#Region "Library Routines"
    Private Function IndexWith() As String
        Dim sWth As String = ""
        Dim sCm As String = ""

        If iFillFactor > 0 Then
            sWth = "fillfactor = " & iFillFactor
            sCm = ", "
        End If
        If bNoRecompute Then
            sWth &= sCm & "statistics_norecompute = on"
            sCm = ", "
        End If
        If bPadIndex Then
            sWth &= sCm & "pad_index = on"
            sCm = ", "
        End If
        If bIgnoreDuplicates Then
            sWth &= sCm & "ignore_dup_key = on"
            sCm = ", "
        End If
        If Not bRowLocking Then
            sWth &= sCm & "allow_row_locks = off"
            sCm = ", "
        End If
        If Not bPageLocking Then
            sWth &= sCm & "allow_page_locks = off"
            sCm = ", "
        End If
        If sWth <> "" Then
            sWth = "with (" & sWth & ")"
        End If
        Return sWth
    End Function

    Private Function Bool2Int(ByVal b As Boolean) As String
        If b Then
            Return "1"
        Else
            Return "0"
        End If
    End Function
#End Region

End Class

Public Class TableIndexes
    Inherits CollectionBase

    Private sClustered As String = ""
    Private sPrimary As String = ""

#Region "Properties"
    Public ReadOnly Property Clustered() As String
        Get
            Clustered = sClustered
        End Get
    End Property

    Public ReadOnly Property PrimaryKey() As String
        Get
            PrimaryKey = sPrimary
        End Get
    End Property

    Default Public Overloads ReadOnly Property Item(ByVal IndexName As String) As TableIndex
        Get
            For Each ti As TableIndex In Me
                If ti.Name = IndexName Then
                    Return ti
                End If
            Next
            Return Nothing
        End Get
    End Property
#End Region

#Region "Methods"
    Public Function Add(ByVal tabndx As TableIndex) As Integer
        If tabndx.Clustered Then
            sClustered = tabndx.Name
        End If
        If tabndx.PrimaryKey Then
            sPrimary = tabndx.Name
        End If
        Return List.Add(tabndx)
    End Function

    Public Function XMLText(ByVal sTab As String, ByVal opt As ScriptOptions) As String
        Dim ss As String = ""
        Dim sp As String = ""
        Dim sOut As String = ""
        Dim cNDX As TableIndex

        For Each cNDX In Me
            If cNDX.PrimaryKey Then
                sOut &= cNDX.PrimaryKeyXML(sTab, opt)
            Else
                ss &= cNDX.XMLText(sTab & "  ")
            End If
        Next
        If ss <> "" Then
            sOut &= sTab & "<indexes>" & vbCrLf
            sOut &= ss
            sOut &= sTab & "</indexes>" & vbCrLf
        End If
        Return sOut
    End Function
#End Region
End Class

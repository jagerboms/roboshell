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

Public Class TableColumn
    Private sType As String = ""
    Private iLength As Integer = 0
    Private iPrecision As Integer = 0
    Private iScale As Integer = 0
    Private sCollation As String = ""
    Private bANSIPadded As Boolean = True
    Private bDefaultNamed As Boolean = False
    Private sqllib As New sql

#Region "Properties"
    Public Name As String
    Public Index As Integer
    Public Nullable As String
    Public RowGuid As Boolean = False
    Public Identity As Boolean = False
    Public Seed As Integer = 1
    Public Increment As Integer = 1
    Public DefaultName As String
    Public DefaultValue As String
    Public Replicated As Boolean = False
    Public Computed As String = ""
    Public Persisted As Boolean = False
    Public XMLDocument As Boolean = False
    Public XMLSchema As String = ""
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

    Public Property DefaultNamed() As Boolean
        Get
            DefaultNamed = bDefaultNamed
        End Get
        Set(ByVal dsn As Boolean)
            bDefaultNamed = dsn
        End Set
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
    ' xml

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
                    If iPrecision <> 53 And iPrecision <> 0 Then
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
                        If XMLSchema <> "" Then
                            s &= XMLSchema & "."
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

    Public Function Text(ByVal sDefCollation As String, ByVal opt As ScriptOptions) As String
        Dim sOut As String

        sOut = QuotedName & " "
        If Computed = "" Then
            sOut &= TypeText
            If Identity Then
                sOut &= " identity(" & Seed & "," & Increment & ")"
                If Replicated Then
                    sOut &= " not for replication"
                End If
            End If

            If RowGuid Then
                sOut &= " rowguidcol"
            End If

            If opt.CollationShow And Collation <> "" Then
                If Collation = sDefCollation Then
                    sOut &= " collate database_default"
                Else
                    sOut &= " collate " & Collation
                End If
            End If

            If Nullable = "N" Then
                sOut &= " not"
            End If
            sOut &= " null"
        Else
            sOut &= "as " & Computed
            If Persisted Then
                sOut &= " persisted"
                If Nullable = "N" Then
                    sOut &= " not null"
                End If
            End If
        End If

        If DefaultName <> "" Then
            If opt.DefaultShowName And Not bDefaultNamed Then
                sOut &= " constraint " & QuotedDefaultName
            End If
            sOut &= " default ("
            If opt.DefaultFix Then
                sOut &= sqllib.FixDefaultText(DefaultValue)
            Else
                sOut &= DefaultValue
            End If
            sOut &= ")"
        End If

        Return sOut
    End Function

    Public Function XMLText(ByVal sTab As String, ByVal sDefCollation As String, ByVal opt As ScriptOptions) As String
        Dim sOut As String
        Dim s As String

        sOut = sTab & "<column name='" & Name & "'"
        If Computed = "" Then
            sOut &= " type='" & Type & "'"
            If Type = "xml" Then
                If XMLCollection <> "" Then
                    If XMLDocument Then
                        sOut &= " document='Y'"
                    Else
                        sOut &= " content='Y'"
                    End If
                    If XMLSchema <> "" Then
                        sOut &= " xmlschema='" & XMLSchema & "'"
                    End If
                    sOut &= " collection='" & XMLCollection & "'"
                End If
            Else
                If Length = -1 Then
                    sOut &= " length='max'"
                ElseIf Length > 0 Then
                    sOut &= " length='" & Length & "'"
                End If
                If Precision > 0 Then
                    sOut &= " precision='" & Precision & "'"
                    sOut &= " scale='" & Scale & "'"
                End If
            End If
                sOut &= " allownulls='" & Nullable & "'"
                If Identity Then
                    sOut &= " seed='" & Seed & "' increment='" & Increment & "'"
                    If Replicated Then
                        sOut &= " replication='N'"
                    End If
                End If
                If RowGuid Then
                    sOut &= " rowguid='Y'"
                End If
                If ANSIPadded = "N" Then
                    sOut &= " ansipadded='N'"
                End If
                If opt.CollationShow And Collation <> "" Then
                    If Collation = sDefCollation Then
                        sOut &= " collation='database_default'"
                    Else
                        sOut &= " collation='" & Collation & "'"
                    End If
                End If
                If DefaultName <> "" Then
                    sOut &= ">" & vbCrLf
                    sOut &= sTab & "  <default"
                    If opt.DefaultShowName And Not bDefaultNamed Then
                        sOut &= " name='" & DefaultName & "'"
                    End If
                    s = DefaultValue
                    If Mid(s, 1, 1) = "(" And Right(s, 1) = ")" Then
                        s = Mid(s, 2, Len(s) - 2)
                    End If
                    If opt.DefaultFix Then
                        s = sqllib.FixDefaultText(s)
                    End If
                    sOut &= "><![CDATA[" & s & "]]></default>" & vbCrLf
                    sOut &= sTab & "</column>"
                Else
                    sOut &= " />"
                End If
        Else
                sOut &= " allownulls='" & Nullable & "'"
                If ANSIPadded = "N" Then
                    sOut &= " ansipadded='N'"
                End If
                If opt.CollationShow And Collation <> "" Then
                    If Collation = sDefCollation Then
                        sOut &= " collation='database_default'"
                    Else
                        sOut &= " collation='" & Collation & "'"
                    End If
                End If
                If Persisted Then
                    sOut &= " persisted='Y'"
                End If
                sOut &= ">" & vbCrLf
                sOut &= sTab & "  <formula><![CDATA[" & Computed & "]]></formula>" & vbCrLf
                sOut &= sTab & "</column>"
        End If

        Return sOut
    End Function
#End Region

End Class

Public Class TableColumns
    Inherits CollectionBase

    Private sIdentity As String = ""
    Private sqllib As New sql

#Region "Properties"
    Default Public Overloads ReadOnly Property Item(ByVal Index As Integer) As TableColumn
        Get
            Return CType(List.Item(Index), TableColumn)
        End Get
    End Property

    Default Public Overloads ReadOnly Property Item(ByVal Name As String) As TableColumn
        Get
            For Each ic As TableColumn In Me
                If ic.Name = Name Then
                    Return ic
                End If
            Next
            Return Nothing
        End Get
    End Property

    Public ReadOnly Property Identity() As String
        Get
            Identity = sIdentity
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub AddColumn( _
    ByVal sName As String, _
    ByVal sType As String, _
    ByVal iLength As Integer, _
    ByVal iPrecision As Integer, _
    ByVal iScale As Integer, _
    ByVal bNullable As String, _
    ByVal sDefaultName As String, _
    ByVal sDefaultValue As String, _
    ByVal bDefaultNamed As Boolean, _
    ByVal sCollation As String, _
    ByVal sANSIPadded As String)

        Dim parm As New TableColumn

        With parm
            .Name = sName
            .Type = sType
            .Length = iLength
            .Precision = iPrecision
            .Scale = iScale
            .Nullable = bNullable
            .DefaultName = sDefaultName
            .DefaultValue = sDefaultValue
            .DefaultNamed = bDefaultNamed
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
        ByVal bDefaultNamed As Boolean, _
        ByVal sCollation As String, _
        ByVal sANSIPadded As String)

        Dim parm As New TableColumn

        With parm
            .Name = sName
            .Type = "xml"
            .XMLSchema = sqllib.QuoteIdentifier(oXMLSchema)
            .XMLCollection = sqllib.QuoteIdentifier(oXMLCollection)
            .XMLDocument = sqllib.GetBit(oXMLDoc, False)
            .Nullable = bNullable
            .DefaultName = sDefaultName
            .DefaultValue = sDefaultValue
            .DefaultNamed = bDefaultNamed
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
        ByVal iLength As Integer, _
        ByVal iPrecision As Integer, _
        ByVal iScale As Integer, _
        ByVal bNullable As String, _
        ByVal sDefaultName As String, _
        ByVal sDefaultValue As String, _
        ByVal bDefaultNamed As Boolean, _
        ByVal sANSIPadded As String)

        Dim parm As New TableColumn

        With parm
            .Name = sName
            .Type = sType
            .Length = iLength
            .Precision = iPrecision
            .Scale = iScale
            .Nullable = bNullable
            .DefaultName = sDefaultName
            .DefaultValue = sDefaultValue
            .DefaultNamed = bDefaultNamed
            .RowGuid = True
            .ANSIPadded = sANSIPadded
        End With

        AddColumn(parm)
    End Sub

    Public Sub AddIdentityColumn( _
        ByVal sName As String, _
        ByVal sType As String, _
        ByVal iLength As Integer, _
        ByVal iPrecision As Integer, _
        ByVal iScale As Integer, _
        ByVal bNullable As String, _
        ByVal iSeed As Integer, _
        ByVal iIncr As Integer, _
        ByVal bRepl As Boolean)

        Dim parm As New TableColumn

        With parm
            .Name = sName
            .Type = sType
            .Length = iLength
            .Precision = iPrecision
            .Scale = iScale
            .Nullable = bNullable
            .Identity = True
            .Seed = iSeed
            .Increment = iIncr
            .Replicated = bRepl
        End With

        AddColumn(parm)
    End Sub

    Public Function XMLText(ByVal sTab As String, ByVal sDefCollation As String, _
                                    ByVal opt As ScriptOptions) As String
        Dim sOut As String
        Dim tc As TableColumn

        sOut = sTab & "<columns>" & vbCrLf
        For Each tc In Me
            sOut &= tc.XMLText(sTab & "  ", sDefCollation, opt) & vbCrLf
        Next
        sOut &= sTab & "</columns>" & vbCrLf
        Return sOut
    End Function
#End Region

#Region "Library Routines"
    Private Sub AddColumn(ByVal tc As TableColumn)
        If tc.Identity Then
            sIdentity = tc.Name
        End If
        tc.Index = List.Count
        List.Add(tc)
    End Sub
#End Region

End Class

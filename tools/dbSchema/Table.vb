Option Explicit On
Option Strict On

' CREATE TABLE
'     [ [ database_name . ] [ schema_name . ] table_name
'         ( { <column_definition> | <computed_column_definition> }
'         [ <table_constraint> ] [ ,...n ] )
'     [ ON { filegroup
'          | partition_scheme_name ( partition_column_name )
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
'             | partition_scheme_name ( partition_column_name )
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
'             | partition_scheme_name ( partition_column_name ) | "default" } ]
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
'             | partition_scheme_name (partition_column_name) | "default" } ]
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

Imports System
Imports System.Data.SqlClient
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

Public Class TableDefn

    Private slib As New sql
    Private PreLoad As Integer = -1
    Private bConsName As Boolean = True
    Private fixdef As Boolean = False
    Private bCollation As Boolean = False

    Private sTable As String
    Private qTable As String
    Private sSchema As String = "dbo"
    Private qSchema As String = "dbo"
    Private sDefCollation As String

    Private fg As New FileGroup
    Private cIndexes As New TableIndexes
    Private cColumns As New TableColumns
    Private cFKeys As New ForeignKeys
    Private cCheckC As New CheckConstraints

    Private dtPerms As DataTable

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
            Return cIndexes.PrimaryKey
        End Get
    End Property

    Public ReadOnly Property State() As Integer
        Get
            Return PreLoad
        End Get
    End Property

    Public ReadOnly Property IKeys() As TableIndexes
        Get
            Return cIndexes
        End Get
    End Property

    Public ReadOnly Property FKeys() As ForeignKeys
        Get
            Return cFKeys
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
            Return cColumns.Identity
        End Get
    End Property

    Public ReadOnly Property hasIdentity() As Boolean
        Get
            If cColumns.Identity = "" Then
                Return False
            End If
            Return True
        End Get
    End Property

    Public ReadOnly Property Column(ByVal index As String) As TableColumn
        Get
            Try
                Return cColumns.Item(index)
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
                    For Each tc As TableColumn In cColumns
                        i = tc.Index + 1
                        If (j And CInt(2 ^ i)) <> 0 Then
                            s &= sC & tc.Name
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

    Public ReadOnly Property XML() As String
        Get
            Dim sOut As String
            Dim s As String
            Dim ss As String
            Dim st As String
            Dim tc As TableColumn
            Dim cNDX As TableIndex
            Dim cFK As ForeignKey
            Dim cCC As CheckConstraint

            sOut = "<?xml version='1.0'?>" & vbCrLf
            sOut &= "<sqldef>" & vbCrLf
            sOut &= "  <table name='" & sTable & "' owner='" & sSchema & "'"
            sOut &= fg.TableXML
            sOut &= fg.TextXML
            sOut &= ">" & vbCrLf

            sOut &= "    <columns>" & vbCrLf
            For Each tc In cColumns
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

            s = cIndexes.PrimaryKey
            If s <> "" Then
                cNDX = cIndexes.Item(s)
                If Not cNDX Is Nothing Then
                    sOut &= cNDX.PrimaryKeyXML("    ", bConsName)
                End If
            End If

            ss = ""
            For Each cNDX In cIndexes
                If Not cNDX.PrimaryKey Then
                    ss &= cNDX.IndexXML("      ")
                End If
            Next
            If ss <> "" Then
                sOut &= "    <indexes>" & vbCrLf
                sOut &= ss
                sOut &= "    </indexes>" & vbCrLf
            End If

            ss = ""
            For Each cCC In cCheckC
                ss &= cCC.CheckXML("      ", bConsName)
            Next
            If ss <> "" Then
                sOut &= "    <constraints>" & vbCrLf
                sOut &= ss
                sOut &= "    </constraints>" & vbCrLf
            End If

            ss = ""
            For Each cFK In cFKeys
                ss &= cFK.ForeignKeyXML("      ")
            Next
            If ss <> "" Then
                sOut &= "    <foreignkeys>" & vbCrLf
                sOut &= ss
                sOut &= "    </foreignkeys>" & vbCrLf
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
        Dim bInc As Boolean
        Dim sdn As String
        Dim sdv As String
        Dim sName As String
        Dim sType As String
        Dim sNull As String
        Dim sColl As String
        Dim sFormula As String
        Dim bPersist As Boolean
        Dim sAP As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim i As Integer
        Dim j As Integer
        Dim iSeed As Integer
        Dim iIncr As Integer
        Dim bRepl As Boolean = False
        Dim cNdx As TableIndex
        Dim cFK As ForeignKey
        Dim cCC As CheckConstraint

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
            sDefCollation = slib.GetString(dr("DefCollation"))
            fg.DefaultGroup = slib.GetString(dr("DefFileGroup"))
            fg.TableGroup = slib.GetString(dr("DataFileGroup"))
            fg.TextGroup = slib.GetString(dr("TextFileGroup"))
            fg.PartitionScheme = slib.GetString(dr("PartitionScheme"))
            fg.SchemeColumn = slib.GetString(dr("SchemeColumn"))
        End If

        For Each dr In dt.Rows        ' Columns
            sName = sqllib.GetString(dr("COLUMN_NAME"))
            sType = sqllib.GetString(dr("DATA_TYPE"))
            sNull = Mid(sqllib.GetString(dr("IS_NULLABLE")), 1, 1)
            sdn = sqllib.GetString(dr("DEFAULT_NAME"))
            sdv = FixDefaultText(sqllib.GetString(dr("DEFAULT_TEXT")))
            sColl = sqllib.GetString(dr("COLLATION_NAME"))
            sAP = Mid(sqllib.GetString(dr("ANSIPadded")), 1, 1)

            If sqllib.GetBit(dr("is_identity"), False) Then
                iSeed = slib.GetInteger(dr("IdentitySeed"), 1)
                iIncr = slib.GetInteger(dr("IdentityIncrement"), 1)
                bRepl = slib.GetBit(dr("IdentityReplicated"), False)
                cColumns.AddIdentityColumn(sName, sType, dr("CHARACTER_MAXIMUM_LENGTH"), _
                    dr.Item("NUMERIC_PRECISION"), dr("NUMERIC_SCALE"), sNull, _
                    iSeed, iIncr, sdn, sdv, sColl, sAP, bRepl)
            Else
                s = sqllib.GetString(dr("ROWGUID"))
                If s = "NO" Then
                    sAP = Mid(sqllib.GetString(dr("ANSIPadded")), 1, 1)
                    sFormula = sqllib.CleanConstraint(sqllib.GetString(dr("Computed")))
                    If sFormula = "" Then
                        If LCase(sType) = "xml" Then
                            cColumns.AddXMLColumn(sName, dr("xmlschema"), _
                                dr("xmlcollection"), dr("is_xml_document"), sNull, _
                                sdn, sdv, sColl, sAP)
                        Else
                            cColumns.AddColumn(sName, sType, dr("CHARACTER_MAXIMUM_LENGTH"), _
                                dr.Item("NUMERIC_PRECISION"), dr("NUMERIC_SCALE"), sNull, _
                                sdn, sdv, sColl, sAP)
                        End If
                    Else
                        If sqllib.GetString(dr("Persisted")) = "NO" Then
                            bPersist = False
                        Else
                            bPersist = True
                        End If
                        cColumns.AddComputedColumn(sName, sFormula, sNull, bPersist)
                    End If
                Else
                    cColumns.AddRowGuidColumn(sName, sType, dr("CHARACTER_MAXIMUM_LENGTH"), _
                        dr.Item("NUMERIC_PRECISION"), dr("NUMERIC_SCALE"), sNull, _
                        sdn, sdv, sAP)
                End If
            End If
        Next

        dt = slib.TableIndexes(qTable, qSchema)
        For Each dr In dt.Rows
            s = sqllib.GetString(dr("name"))

            cNdx = cIndexes(s)
            If cNdx Is Nothing Then
                cNdx = New TableIndex(sSchema, sTable, s)
                cNdx.FillFactor = slib.GetInteger(dr("FILL_FACTOR"), 0)
                i = slib.GetInteger(dr("type"), 2)
                If i = 1 Then
                    cNdx.Clustered = True
                End If
                If slib.GetBit(dr("is_primary_key"), False) Then
                    cNdx.PrimaryKey = True
                End If
                If slib.GetBit(dr("is_unique"), False) Then
                    cNdx.Unique = True
                End If
                If slib.GetBit(dr("no_recompute"), False) Then
                    cNdx.NoRecompute = True
                End If
                If slib.GetBit(dr("PAD_INDEX"), False) Then
                    cNdx.PadIndex = True
                End If
                If slib.GetBit(dr("IGNORE_DUP_KEY"), False) Then
                    cNdx.IgnoreDuplicates = True
                End If
                If Not slib.GetBit(dr("ALLOW_ROW_LOCKS"), True) Then
                    cNdx.RowLocking = False
                End If
                If Not slib.GetBit(dr("ALLOW_PAGE_LOCKS"), True) Then
                    cNdx.PageLocking = False
                End If
                cNdx.IndexFileGroup.DefaultGroup = fg.DefaultGroup
                cNdx.IndexFileGroup.TableGroup = fg.TableGroup
                cNdx.IndexFileGroup.TextGroup = fg.TextGroup
                cNdx.IndexFileGroup.IndexGroup = slib.GetString(dr("filegroup"))

                cIndexes.Add(cNdx)
            End If
            s = sqllib.GetString(dr("ColumnName"))
            j = sqllib.GetInteger(dr("key_ordinal"), 0)
            b = sqllib.GetBit(dr("is_descending_key"), False)
            bInc = sqllib.GetBit(dr("is_included_column"), False)
            i = sqllib.GetInteger(dr("partition_ordinal"), 0)
            If i > 0 Then
                cNdx.IndexFileGroup.SchemeColumn = s
            End If
            cNdx.Columns.Add(s, j, b, bInc, i)
        Next

        dt = slib.TableCheckConstraints(qTable, qSchema)
        For Each r As DataRow In dt.Rows
            s = slib.GetString(r("ConstraintName"))
            cCC = New CheckConstraint(sSchema, sTable, s)
            cCC.Definition = slib.GetString(r("Definition"))
            If slib.GetBit(r("Replicated"), False) Then
                cCC.Replicated = True
            End If
            cCC.ColumnName = slib.GetString(r("ColumnName"))
            cCC.SystemName = slib.GetBit(r("SystemName"), False)
            cCC.IsSystem = slib.GetBit(r("IsSystem"), False)
            cCheckC.Add(cCC)
        Next

        dt = slib.TableFKeys(qTable, qSchema)
        For Each r As DataRow In dt.Rows
            s = slib.GetString(r("ConstraintName"))
            cFK = cFKeys(s)
            If cFK Is Nothing Then
                cFK = New ForeignKey(sSchema, sTable, s)
                cFK.LinkedSchema = slib.GetString(r("LinkedSchema"))
                cFK.LinkedTable = slib.GetString(r("LinkedTable"))
                cFK.MatchOption = slib.GetString(r("MATCH_OPTION"))
                cFK.DeleteOption = slib.GetString(r("DELETE_RULE"))
                cFK.UpdateOption = slib.GetString(r("UPDATE_RULE"))
                If slib.GetBit(r("Replicated"), False) Then
                    cFK.Replicated = True
                End If
                cFKeys.Add(cFK)
            End If
            i = slib.GetInteger(r("Sequence"), -1)
            sName = slib.GetString(r("ColumnName"))
            s = slib.GetString(r("LinkedColumn"))

            cFK.Columns.Add(sName, s, i)
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

    Public Function DataScript(ByVal sFilter As String) As String
        Dim tc As TableColumn
        Dim dt As DataTable
        Dim sOut As String = ""
        Dim sHead As String
        Dim sTail As String = ""
        Dim s As String
        Dim qN As String
        Dim i As Integer
        Dim ss As String = ""
        Dim cNDX As TableIndex = Nothing

        If cColumns.Identity <> "" Then
            sOut &= "set identity_insert " & qSchema & "." & qTable & " on" & vbCrLf
            sOut &= vbCrLf
        End If

        sHead = "insert into " & qSchema & "." & qTable & vbCrLf
        sHead &= "(" & vbCrLf
        s = "    "
        For Each tc In cColumns
            sHead &= s & tc.QuotedName
            s = ", "
        Next
        sHead &= vbCrLf
        sHead &= ")" & vbCrLf
        s = "select  x."
        For Each tc In cColumns
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
                For Each tc In cColumns
                    sOut &= s & tc.DataFormat(r(tc.Name)) & " " & tc.QuotedName & vbCrLf
                    s = "           ,"
                Next
                sTail = ") x" & vbCrLf
                sTail &= "left join " & qSchema & "." & qTable & " a" & vbCrLf

                s = cIndexes.PrimaryKey
                If s <> "" Then
                    cNDX = cIndexes.Item(s)
                End If

                If cNDX Is Nothing Then
                    Return ""
                End If

                s = "on      a."
                For Each ic As IndexColumn In cNDX.Columns
                    qN = slib.QuoteIdentifier(ic.Name)
                    If ss = "" Then ss = qN
                    sTail &= s & qN & " = x." & qN & vbCrLf
                    s = "and     a."
                Next

                sTail &= "where   a." & ss & " is null" & vbCrLf
                sTail &= "go" & vbCrLf & vbCrLf

                i += 1
            Else
                s = "    union select "
                For Each tc In cColumns
                    sOut &= s & tc.DataFormat(r(tc.Name))
                    s = ", "
                Next
                sOut &= vbCrLf
                i += 1
            End If
            If i = 100 Then i = 0
        Next

        sOut &= sTail
        If cColumns.Identity <> "" Then
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

        For Each tc In cColumns
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

        For Each tc In cColumns
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

        s = cIndexes.PrimaryKey
        If s <> "" Then
            Dim c As TableIndex
            c = cIndexes.Item(s)
            If Not c Is Nothing Then
                sOut &= c.PrimaryKeyText(sTab, bConsName)
            End If
        End If

        For Each ck As CheckConstraint In cCheckC
            sOut &= ck.CheckText(sTab & "   ", bConsName)
        Next
        sOut &= sTab & ")"
        sOut &= fg.TableText
        sOut &= fg.TextText
        sOut &= vbCrLf

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

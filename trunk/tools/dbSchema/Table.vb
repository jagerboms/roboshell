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
    Private cPerms As New TablePermissions

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

    Public ReadOnly Property QuotedName() As String
        Get
            QuotedName = qTable
        End Get
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

    Public ReadOnly Property QuotedSchema() As String
        Get
            QuotedSchema = qSchema
        End Get
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

    Public ReadOnly Property Permissions() As TablePermissions
        Get
            Return cPerms
        End Get
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

    Public ReadOnly Property AllColumns() As TableColumns
        Get
            Try
                Return cColumns
            Catch
                Return Nothing
            End Try
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub New()
        PreLoad = 0
    End Sub

    Public Sub New(ByVal sTableName As String, ByVal sqllib As sql)
        slib = sqllib
        LoadTable(sTableName, "dbo")
    End Sub

    Public Sub New(ByVal sTableName As String, ByVal Schema As String, _
                   ByVal sqllib As sql)
        slib = sqllib
        LoadTable(sTableName, Schema)
    End Sub

    Private Sub LoadTable(ByVal sTableName As String, ByVal Sch As String)
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
        Dim cPm As TablePermission

        sSchema = Sch
        qSchema = slib.QuoteIdentifier(Sch)
        PreLoad = 2

        dt = slib.TableColumns(slib.QuoteIdentifier(sTableName), qSchema)
        If dt.Rows.Count = 0 Then
            PreLoad = 3
            Return
        End If

        sTable = slib.GetString(dt.Rows(0).Item("TableName"))
        qTable = slib.QuoteIdentifier(sTable)
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
            sName = slib.GetString(dr("COLUMN_NAME"))
            sType = slib.GetString(dr("DATA_TYPE"))
            sNull = Mid(slib.GetString(dr("IS_NULLABLE")), 1, 1)
            sdn = slib.GetString(dr("DEFAULT_NAME"))
            sdv = slib.CleanConstraint((slib.GetString(dr("DEFAULT_TEXT"))))
            sColl = slib.GetString(dr("COLLATION_NAME"))
            sAP = Mid(slib.GetString(dr("ANSIPadded")), 1, 1)

            If slib.GetBit(dr("is_identity"), False) Then
                iSeed = slib.GetInteger(dr("IdentitySeed"), 1)
                iIncr = slib.GetInteger(dr("IdentityIncrement"), 1)
                bRepl = slib.GetBit(dr("IdentityReplicated"), False)
                cColumns.AddIdentityColumn(sName, sType, dr("CHARACTER_MAXIMUM_LENGTH"), _
                    dr.Item("NUMERIC_PRECISION"), dr("NUMERIC_SCALE"), sNull, _
                    iSeed, iIncr, sdn, sdv, sColl, sAP, bRepl)
            Else
                s = slib.GetString(dr("ROWGUID"))
                If s = "NO" Then
                    sAP = Mid(slib.GetString(dr("ANSIPadded")), 1, 1)
                    sFormula = slib.CleanConstraint(slib.GetString(dr("Computed")))
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
                        If slib.GetString(dr("Persisted")) = "NO" Then
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
            s = slib.GetString(dr("name"))

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
            s = slib.GetString(dr("ColumnName"))
            j = slib.GetInteger(dr("key_ordinal"), 0)
            b = slib.GetBit(dr("is_descending_key"), False)
            bInc = slib.GetBit(dr("is_included_column"), False)
            i = slib.GetInteger(dr("partition_ordinal"), 0)
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

        dt = slib.TablePermissions(qTable, qSchema)
        For Each r As DataRow In dt.Rows
            cPm = cPerms.Add(slib.GetString(r("Grantee")), _
                    qSchema & "." & qTable, _
                    slib.GetString(r("Permissions")))

            s = slib.GetString(r("State"))
            If s = "GRANT_WITH_GRANT_OPTION" Then
                cPm.GrantOption = True
            ElseIf s = "DENY" Then
                cPm.Deny = True
            End If
            j = slib.GetInteger(dr("Columns"), 0)
            If j > 1 Then
                For Each tc As TableColumn In cColumns
                    i = tc.Index + 1
                    If (j And CInt(2 ^ i)) <> 0 Then
                        cPm.AddColumn(tc.Name)
                    End If
                Next
            End If
        Next

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

    Public Function XML(ByVal opt As ScriptOptions) As String
        Dim sOut As String

        sOut = "<?xml version='1.0'?>" & vbCrLf
        sOut &= "<sqldef>" & vbCrLf
        sOut &= "  <table name='" & sTable & "' owner='" & sSchema & "'"
        sOut &= fg.TableXML
        sOut &= fg.TextXML
        sOut &= ">" & vbCrLf
        sOut &= cColumns.XMLText("    ", sDefCollation, opt)
        sOut &= cIndexes.XMLText("    ", opt)
        sOut &= cCheckC.XMLText("    ", opt)
        sOut &= cFKeys.XMLText("    ")
        sOut &= "  </table>" & vbCrLf
        sOut &= "</sqldef>" & vbCrLf
        XML = sOut
    End Function

    Public Function TableText(ByVal opt As ScriptOptions) As String
        Return CreateTable(False, opt)
    End Function

    Public Function FullTableText(ByVal opt As ScriptOptions) As String
        Return CreateTable(True, opt)
    End Function
#End Region

#Region "private functions"
    Private Function CreateTable(ByVal bFull As Boolean, ByVal opt As ScriptOptions) As String
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

                If opt.CollationShow And tc.Collation <> "" Then
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
                If opt.DefaultShowName Then
                    sOut &= " constraint " & tc.QuotedDefaultName
                End If
                sOut &= " default ("
                If opt.DefaultFix Then
                    sOut &= slib.FixDefaultText(tc.DefaultValue)
                Else
                    sOut &= tc.DefaultValue
                End If
                sOut &= ")"
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
                sOut &= c.PrimaryKeyText(sTab, opt)
            End If
        End If

        For Each ck As CheckConstraint In cCheckC
            sOut &= ck.Text(sTab & "   ", opt)
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
#End Region
End Class

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

    Public Sub New(ByVal sqllib As sql, ByVal XML As Xml.XmlElement)
        slib = sqllib
        LoadXML(XML)
    End Sub

    Private Sub LoadTable(ByVal sTableName As String, ByVal Sch As String)
        Dim s As String = "a"
        Dim b As Boolean = False
        Dim bInc As Boolean
        Dim sdn As String
        Dim sdv As String
        Dim bn As Boolean
        Dim sName As String
        Dim sType As String
        Dim sNull As String
        Dim iLen As Integer = 0
        Dim iPrc As Integer = 0
        Dim iScl As Integer = 0
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
            iLen = slib.GetInteger(dr("CHARACTER_MAXIMUM_LENGTH"), 0)
            iPrc = slib.GetInteger(dr("NUMERIC_PRECISION"), 0)
            iScl = slib.GetInteger(dr("NUMERIC_SCALE"), 0)
            sNull = Mid(slib.GetString(dr("IS_NULLABLE")), 1, 1)
            sdn = slib.GetString(dr("DEFAULT_NAME"))
            sdv = slib.CleanConstraint((slib.GetString(dr("DEFAULT_TEXT"))))
            bn = slib.GetBit(dr("DefaultNamed"), False)
            sColl = slib.GetString(dr("COLLATION_NAME"))
            sAP = Mid(slib.GetString(dr("ANSIPadded")), 1, 1)

            If slib.GetBit(dr("is_identity"), False) Then
                iSeed = slib.GetInteger(dr("IdentitySeed"), 1)
                iIncr = slib.GetInteger(dr("IdentityIncrement"), 1)
                bRepl = slib.GetBit(dr("IdentityReplicated"), False)
                cColumns.AddIdentityColumn(sName, sType, iLen, iPrc, iScl, _
                    sNull, iSeed, iIncr, bRepl)
            Else
                s = slib.GetString(dr("ROWGUID"))
                If s = "NO" Then
                    sAP = Mid(slib.GetString(dr("ANSIPadded")), 1, 1)
                    sFormula = slib.CleanConstraint(slib.GetString(dr("Computed")))
                    If sFormula = "" Then
                        If LCase(sType) = "xml" Then
                            cColumns.AddXMLColumn(sName, dr("xmlschema"), _
                                dr("xmlcollection"), dr("is_xml_document"), sNull, _
                                sdn, sdv, bn, sColl, sAP)
                        Else
                            cColumns.AddColumn(sName, sType, iLen, iPrc, iScl, sNull, _
                                sdn, sdv, bn, sColl, sAP)
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
                    cColumns.AddRowGuidColumn(sName, sType, iLen, iPrc, iScl, sNull, _
                        sdn, sdv, bn, sAP)
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
                    If slib.GetBit(dr("PrimaryKeyNamed"), False) Then
                        cNdx.PrimaryKeyNamed = True
                    End If
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
        For Each dr In dt.Rows
            s = slib.GetString(dr("ConstraintName"))
            cCC = New CheckConstraint(sSchema, sTable, s)
            cCC.Definition = slib.GetString(dr("Definition"))
            If slib.GetBit(dr("Replicated"), False) Then
                cCC.Replicated = True
            End If
            cCC.ColumnName = slib.GetString(dr("ColumnName"))
            cCC.SystemName = slib.GetBit(dr("SystemName"), False)
            cCC.IsSystem = slib.GetBit(dr("IsSystem"), False)
            cCheckC.Add(cCC)
        Next

        dt = slib.TableFKeys(qTable, qSchema)
        For Each dr In dt.Rows
            s = slib.GetString(dr("ConstraintName"))
            cFK = cFKeys(s)
            If cFK Is Nothing Then
                cFK = New ForeignKey(sSchema, sTable, s)
                cFK.LinkedSchema = slib.GetString(dr("LinkedSchema"))
                cFK.LinkedTable = slib.GetString(dr("LinkedTable"))
                cFK.MatchOption = slib.GetString(dr("MATCH_OPTION"))
                cFK.DeleteOption = slib.GetString(dr("DELETE_RULE"))
                cFK.UpdateOption = slib.GetString(dr("UPDATE_RULE"))
                If slib.GetBit(dr("Replicated"), False) Then
                    cFK.Replicated = True
                End If
                cFKeys.Add(cFK)
            End If
            i = slib.GetInteger(dr("Sequence"), -1)
            sName = slib.GetString(dr("ColumnName"))
            s = slib.GetString(dr("LinkedColumn"))

            cFK.Columns.Add(sName, s, i)
        Next

        dt = slib.TablePermissions(qTable, qSchema)
        For Each dr In dt.Rows
            cPm = cPerms.Add(slib.GetString(dr("Grantee")), _
                    qSchema & "." & qTable, _
                    slib.GetString(dr("permission_name")))

            s = slib.GetString(dr("State"))
            If s = "GRANT_WITH_GRANT_OPTION" Then
                cPm.GrantOption = True
            ElseIf s = "DENY" Then
                cPm.Deny = True
            End If
            j = slib.GetInteger(dr("columns"), 0)
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

    '<?xml version='1.0'?>
    '<sqldef>
    '  <table name='ttt' owner='dbo'
    '      [filegroup='EXTENSION' | partitionfunction='FUNCTION' partitioncolumn='COL'
    '               textfilegroup='EXTENSION']
    '    <columns>
    '      <column name='id' type='integer' length='max' allownulls='N' seed='1' increment='1' />

    '<column name='Name' type='Type'
    ' document='Y' | content='Y'
    '  collection='XMLCollection'
    '  length='Length'
    '  precision='Precision' scale='Scale'
    '  allownulls='Y'
    '  seed='Seed' increment='Increment' replication='N'
    '  rowguid='Y'
    '  ansipadded='N'
    '  collation='database_default'
    '  persisted='Y'>
    '  <formula><![CDATA[Computed]]></formula>
    '        <default name='DEFAULTNAME'><![CDATA[getdate()]]></default>
    ' </column>
    '      <column name='txt' type='text' length='max' allownulls='Y' />
    '      <column name='dt' type='datetime' length='max' allownulls='N'>
    '        <default><![CDATA[getdate()]]></default>
    '      </column>
    '    </columns>
    '    <primarykey [name='PK'] clustered='Y'
    '       [fillfactor='90']
    '       [[filegroup='EXTENSION'] | [partitionfunction='FUNCTION' partitioncolumn='COL']]
    '       [pad='on']
    '       [dup='on']
    '       [rowlocks='off']
    '       [pagelocks='off'] >
    '      <column name='id' [direction='desc'] />
    '    </primarykey>
    '    <indexes>
    '       <index name='Name'
    '          [clustered='Y']
    '          [[filegroup='EXTENSION'] | [partitionfunction='FUNCTION' partitioncolumn='COL']]
    '          [unique='Y']
    '          [fillfactor='90']
    '          [pad='on']
    '          [dup='on']
    '          [rowlocks='off']
    '          [pagelocks='off'] >
    '          <column name='Name'
    '             [included='Y']
    '             [direction='desc'] /> ...
    '       </index>
    '    </indexes>
    '    <constraints>
    '       <constraint type='check'
    '          [name='Name']
    '          [column='Name']
    '          [replication='N'] >
    '          <![CDATA[col='Y' col='N']]>
    '       </constraint>
    '    </constraints>
    '    <foreignkeys>
    '      <foreignkey name='Name' references='LinkedTable'
    '        [match='MCH']
    '        [ondelete='DEL']
    '        [onupdate='UPD']
    '        [replication='N'] >
    '          <column name='Name' linksto='Linked' />
    '      </foreignkey>
    ''   </foreignkeys>
    '  </table>
    '</sqldef>

    Public Sub LoadXML(ByVal x As Xml.XmlElement)
        Dim b As Boolean
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim s As String
        Dim ele As Xml.XmlElement
        Dim ele0 As Xml.XmlElement
        Dim ele1 As Xml.XmlElement
        Dim att As Xml.XmlAttribute
        Dim sName As String
        Dim sType As String
        Dim sNull As String
        Dim iLen As Integer
        Dim iPrc As Integer
        Dim iScl As Integer
        Dim iSeed As Integer
        Dim iIncr As Integer
        Dim bRepl As Boolean
        Dim sAP As String
        Dim sColl As String
        Dim bRG As Boolean
        Dim sFormula As String
        Dim bPersist As Boolean
        Dim bDoc As Boolean
        Dim sXMLColl As String
        Dim sXMLSch As String
        Dim sdn As String
        Dim sdv As String
        Dim bn As Boolean
        Dim cNdx As TableIndex
        Dim cCC As CheckConstraint
        Dim cFK As ForeignKey

        If x.Name = "table" Then
            For Each att In x.Attributes
                Select Case att.Name
                    Case "name"
                        sTable = att.InnerText
                        qTable = slib.QuoteIdentifier(sTable)

                    Case "owner"
                        sSchema = att.InnerText
                        qSchema = slib.QuoteIdentifier(sSchema)

                    Case "filegroup"
                        fg.TableGroup = att.InnerText

                    Case "textfilegroup"
                        fg.TextGroup = att.InnerText

                    Case "partitionfunction"
                        fg.PartitionScheme = att.InnerText

                    Case "partitioncolumn"
                        fg.SchemeColumn = att.InnerText

                End Select
            Next

            For Each ele In x.ChildNodes
                Select Case ele.Name
                    Case "columns"
                        For Each ele0 In ele.ChildNodes
                            sName = ""
                            sType = "char"
                            sNull = "N"
                            iLen = 0
                            iPrc = 0
                            iScl = 0
                            sAP = "Y"
                            sColl = ""
                            iSeed = 0
                            iIncr = 0
                            bRepl = False
                            bRG = False
                            bPersist = False
                            bDoc = False
                            sXMLColl = ""
                            sXMLSch = ""

                            For Each att In ele0.Attributes
                                Select Case att.Name
                                    Case "name"
                                        sName = att.InnerText

                                    Case "type"
                                        sType = att.InnerText

                                    Case "allownulls"
                                        sNull = att.InnerText

                                    Case "length"
                                        If LCase(att.InnerText) = "max" Then
                                            iLen = -1
                                        Else
                                            iLen = CInt(att.InnerText)
                                        End If

                                    Case "precision"
                                        iPrc = CInt(att.InnerText)

                                    Case "scale"
                                        iScl = CInt(att.InnerText)

                                    Case "ansipadded"
                                        sAP = UCase(Mid(att.InnerText, 1, 1))

                                    Case "collation"
                                        sColl = att.InnerText

                                    Case "seed"
                                        iSeed = CInt(att.InnerText)

                                    Case "increment"
                                        iIncr = CInt(att.InnerText)

                                    Case "replication"
                                        If LCase(Mid(att.InnerText, 1, 1)) = "n" Then
                                            bRepl = True
                                        End If

                                    Case "rowguid"
                                        If LCase(Mid(att.InnerText, 1, 1)) = "y" Then
                                            bRG = True
                                        End If

                                    Case "persisted"
                                        If LCase(Mid(att.InnerText, 1, 1)) = "y" Then
                                            bPersist = True
                                        End If

                                    Case "document"
                                        If LCase(Mid(att.InnerText, 1, 1)) = "y" Then
                                            bDoc = True
                                        End If

                                    Case "content"
                                        If LCase(Mid(att.InnerText, 1, 1)) = "y" Then
                                            bDoc = False
                                        End If

                                    Case "collection"
                                        sXMLColl = att.InnerText

                                    Case "xmlschema"
                                        sXMLSch = att.InnerText

                                End Select
                            Next
                            If sName = "" Then
                                PreLoad = 3
                                Return
                            End If
                            sFormula = ""
                            sdn = ""
                            sdv = ""
                            bn = False
                            For Each ele1 In ele0.ChildNodes
                                Select Case ele1.Name
                                    Case "formula"
                                        sFormula = ele1.InnerText

                                    Case "default"
                                        For Each att In ele1.Attributes
                                            Select Case att.Name
                                                Case "name"
                                                    sdn = att.InnerText
                                            End Select
                                        Next
                                        sdv = ele1.InnerText

                                        If sdn = "" Then
                                            sdn = sName & "Def"
                                            bn = True
                                        End If
                                End Select
                            Next

                            If iSeed > 0 Then
                                cColumns.AddIdentityColumn(sName, sType, iLen, iPrc, _
                                               iScl, sNull, iSeed, iIncr, bRepl)
                            ElseIf bRG Then
                                cColumns.AddRowGuidColumn(sName, sType, iLen, iPrc, _
                                               iScl, sNull, sdn, sdv, bn, sAP)
                            ElseIf sFormula <> "" Then
                                cColumns.AddComputedColumn(sName, sFormula, _
                                               sNull, bPersist)
                            ElseIf LCase(sType) = "xml" Then
                                cColumns.AddXMLColumn(sName, sXMLSch, sXMLColl, bDoc, _
                                               sNull, sdn, sdv, bn, sColl, sAP)
                            Else
                                cColumns.AddColumn(sName, sType, iLen, iPrc, iScl, _
                                               sNull, sdn, sdv, bn, sColl, sAP)
                            End If
                        Next

                    Case "primarykey"
                        sName = ""
                        bn = False
                        For Each att In ele.Attributes
                            Select Case att.Name
                                Case "name"
                                    sName = att.InnerText
                                    Exit For
                            End Select
                        Next
                        If sName = "" Then
                            sName = sTable & "PK"
                            bn = True
                        End If

                        cNdx = New TableIndex(sSchema, sTable, sName)
                        cNdx.PrimaryKey = True
                        cNdx.PrimaryKeyNamed = bn
                        cNdx.IndexFileGroup.DefaultGroup = fg.DefaultGroup
                        cNdx.IndexFileGroup.TableGroup = fg.TableGroup
                        cNdx.IndexFileGroup.TextGroup = fg.TextGroup

                        b = False
                        sColl = ""
                        s = ""
                        For Each att In ele.Attributes
                            Select Case att.Name
                                Case "clustered"
                                    If LCase(Mid(att.InnerText, 1, 1)) = "y" Then
                                        cNdx.Clustered = True
                                    End If

                                Case "fillfactor"
                                    cNdx.FillFactor = CInt(att.InnerText)

                                Case "pad"
                                    If LCase(att.InnerText) = "on" Then
                                        cNdx.PadIndex = True
                                    End If

                                Case "dup"
                                    If LCase(att.InnerText) = "on" Then
                                        cNdx.IgnoreDuplicates = True
                                    End If

                                Case "rowlocks"
                                    If LCase(att.InnerText) = "off" Then
                                        cNdx.RowLocking = False
                                    End If

                                Case "pagelocks"
                                    If LCase(att.InnerText) = "off" Then
                                        cNdx.PageLocking = False
                                    End If

                                Case "filegroup"
                                    cNdx.IndexFileGroup.IndexGroup = att.InnerText

                                Case "partitionfunction"
                                    cNdx.IndexFileGroup.PartitionScheme = att.InnerText

                                Case "partitioncolumn"
                                    s = att.InnerText
                                    cNdx.IndexFileGroup.SchemeColumn = s

                            End Select
                        Next
                        cIndexes.Add(cNdx)
                        i = 1
                        j = 0
                        For Each ele1 In ele.ChildNodes
                            sName = ""
                            b = False
                            For Each att In ele1.Attributes
                                Select Case att.Name
                                    Case "name"
                                        sName = att.InnerText
                                        If sName = s Then
                                            j = 1
                                        End If

                                    Case "direction"
                                        If LCase(att.InnerText) = "desc" Then
                                            b = True
                                        End If

                                End Select
                            Next
                            cNdx.Columns.Add(sName, i, b, False, j)
                            i += 1
                        Next

                    Case "indexes"
                        For Each ele0 In ele.ChildNodes
                            sName = ""
                            s = ""
                            For Each att In ele0.Attributes
                                Select Case att.Name
                                    Case "name"
                                        sName = att.InnerText
                                        Exit For
                                End Select
                            Next
                            cNdx = New TableIndex(sSchema, sTable, sName)
                            cNdx.IndexFileGroup.DefaultGroup = fg.DefaultGroup
                            cNdx.IndexFileGroup.TableGroup = fg.TableGroup
                            cNdx.IndexFileGroup.TextGroup = fg.TextGroup

                            sColl = ""
                            For Each att In ele0.Attributes
                                Select Case att.Name
                                    Case "unique"
                                        If LCase(Mid(att.InnerText, 1, 1)) = "y" Then
                                            cNdx.Unique = True
                                        End If

                                    Case "clustered"
                                        If LCase(Mid(att.InnerText, 1, 1)) = "y" Then
                                            cNdx.Clustered = True
                                        End If

                                    Case "fillfactor"
                                        cNdx.FillFactor = CInt(att.InnerText)

                                    Case "norecompute"
                                        If LCase(Mid(att.InnerText, 1, 1)) = "y" Then
                                            cNdx.NoRecompute = True
                                        End If

                                    Case "pad"
                                        If LCase(att.InnerText) = "on" Then
                                            cNdx.PadIndex = True
                                        End If

                                    Case "dup"
                                        If LCase(att.InnerText) = "on" Then
                                            cNdx.IgnoreDuplicates = True
                                        End If

                                    Case "rowlocks"
                                        If LCase(att.InnerText) = "off" Then
                                            cNdx.RowLocking = False
                                        End If

                                    Case "pagelocks"
                                        If LCase(att.InnerText) = "off" Then
                                            cNdx.PageLocking = False
                                        End If

                                    Case "filegroup"
                                        cNdx.IndexFileGroup.IndexGroup = att.InnerText

                                    Case "partitionfunction"
                                        cNdx.IndexFileGroup.PartitionScheme = att.InnerText

                                    Case "partitioncolumn"
                                        s = att.InnerText
                                        cNdx.IndexFileGroup.SchemeColumn = s

                                End Select
                            Next
                            cIndexes.Add(cNdx)
                            i = 1
                            For Each ele1 In ele0.ChildNodes
                                sName = ""
                                b = False
                                bn = False
                                j = 0
                                For Each att In ele1.Attributes
                                    Select Case att.Name
                                        Case "name"
                                            sName = att.InnerText
                                            If sName = s Then
                                                j = 1
                                            End If

                                        Case "direction"
                                            If LCase(att.InnerText) = "desc" Then
                                                b = True
                                            End If

                                        Case "included"
                                            If LCase(Mid(att.InnerText, 1, 1)) = "y" Then
                                                bn = True
                                            End If

                                    End Select
                                Next
                                If bn Then
                                    k = 0
                                Else
                                    k = i
                                    i += 1
                                End If
                                cNdx.Columns.Add(sName, k, b, bn, j)
                            Next
                        Next

                    Case "constraints"
                        i = 0
                        For Each ele0 In ele.ChildNodes
                            sType = ""
                            sName = ""
                            s = ""
                            bRepl = False
                            For Each att In ele0.Attributes
                                Select Case att.Name
                                    Case "type"
                                        sType = att.InnerText

                                    Case "name"
                                        sName = att.InnerText

                                    Case "column"
                                        s = att.InnerText

                                    Case "replication"
                                        If LCase(Mid(att.InnerText, 1, 1)) = "n" Then
                                            bRepl = True
                                        End If

                                End Select
                            Next
                            If sType = "check" Then
                                b = False
                                If sName = "" Then
                                    sName = "check" & i
                                    i += 1
                                    b = True
                                End If
                                cCC = New CheckConstraint(sSchema, sTable, sName)
                                cCC.Definition = ele0.InnerText
                                cCC.Replicated = bRepl
                                cCC.ColumnName = s
                                cCC.SystemName = b
                                cCheckC.Add(cCC)
                            End If
                        Next

                    Case "foreignkeys"
                        For Each ele0 In ele.ChildNodes
                            sName = ""
                            For Each att In ele0.Attributes
                                Select Case att.Name
                                    Case "name"
                                        sName = att.InnerText
                                End Select
                            Next
                            cFK = New ForeignKey(sSchema, sTable, sName)
                            For Each att In ele0.Attributes
                                Select Case att.Name
                                    Case "references"
                                        cFK.LinkedTable = att.InnerText

                                    Case "referenceschema"
                                        cFK.LinkedSchema = att.InnerText

                                    Case "match"
                                        cFK.MatchOption = att.InnerText

                                    Case "ondelete"
                                        cFK.DeleteOption = att.InnerText

                                    Case "onupdate"
                                        cFK.UpdateOption = att.InnerText

                                    Case "replication"
                                        If LCase(Mid(att.InnerText, 1, 1)) = "n" Then
                                            cFK.Replicated = True
                                        End If

                                End Select
                            Next
                            cFKeys.Add(cFK)

                            i = 1
                            For Each ele1 In ele0.ChildNodes
                                sName = ""
                                s = ""
                                b = False
                                For Each att In ele1.Attributes
                                    Select Case att.Name
                                        Case "name"
                                            sName = att.InnerText

                                        Case "linksto"
                                            s = att.InnerText

                                    End Select
                                Next
                                cFK.Columns.Add(sName, s, i)
                                i += 1
                            Next
                        Next

                End Select
            Next
        End If
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
        If opt.TargetEnvironment = ScriptOptions.TargetEnvironments.PostGres Then
            Return PostGresTableText(False, opt)
        Else
            Return CreateTable(False, opt)
        End If
    End Function

    Public Function FullTableText(ByVal opt As ScriptOptions) As String
        If opt.TargetEnvironment = ScriptOptions.TargetEnvironments.PostGres Then
            Return PostGresTableText(True, opt)
        Else
            Return CreateTable(True, opt)
        End If
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
            sOut = "if object_id(" & slib.QuoteString(sSchema & "." & sTable) & ") is null" & vbCrLf
            sOut &= "begin" & vbCrLf
            sOut &= "    print 'creating " & slib.GetSQLString(sSchema & "." & sTable) & "'" & vbCrLf
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
            sOut &= sTab & "   " & Comma & tc.Text(sDefCollation, opt)
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

    Private Function PostGresTableText(ByVal bFull As Boolean, ByVal opt As ScriptOptions) As String
        Dim sOut As String = ""
        Dim Comma As String
        Dim s As String
        Dim sTab As String
        Dim tc As TableColumn

        sTab = "        "

        If PreLoad = 3 Then
            Return ""
        End If

        sOut = "create function dbo.oOo_ct_" & sTable & "()" & vbCrLf
        sOut &= "returns(void)" & vbCrLf
        sOut &= "as $$" & vbCrLf
        sOut &= "begin" & vbCrLf
        sOut &= "    if not exists Then" & vbCrLf
        sOut &= "    (" & vbCrLf
        sOut &= sTab & "select  'a'" & vbCrLf
        sOut &= sTab & "from(information_schema.tables)" & vbCrLf
        sOut &= sTab & "where   table_name = '" & sTable & "'" & vbCrLf
        sOut &= sTab & "and     table_schema = '" & sSchema & "'" & vbCrLf
        sOut &= "    ) then" & vbCrLf
        sOut &= vbCrLf
        sOut &= sTab & "create table " & qSchema & "." & qTable & vbCrLf
        sOut &= sTab & "(" & vbCrLf
        Comma = " "
        For Each tc In cColumns
            sOut &= sTab & "   " & Comma & tc.Text(sDefCollation, opt) & vbCrLf

            ' sysname, smalldatetime, datetime

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
        sOut &= sTab & ");"
        sOut &= vbCrLf

        If cIndexes.Clustered = "Y" Then
            sOut &= sTab & "alter table " & sSchema & "." & sTable & " cluster on " & s & ";" & vbCrLf
        End If
        sOut &= "    end if;" & vbCrLf
        sOut &= "end;" & vbCrLf
        sOut &= "$$ language plpgsql;" & vbCrLf
        sOut &= vbCrLf
        sOut &= "select dbo.oOo_ct_" & sTable & "();" & vbCrLf
        sOut &= "drop function dbo.oOo_ct_" & sTable & "();" & vbCrLf
        sOut &= vbCrLf

        Return sOut
    End Function
#End Region
End Class

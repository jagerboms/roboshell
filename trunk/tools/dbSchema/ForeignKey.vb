Option Explicit On
Option Strict On

' [ CONSTRAINT constraint_name ]
' {
'   FOREIGN KEY ( column [ ,...n ] )
'   REFERENCES referenced_table_name [ ( ref_column [ ,...n ] ) ]
'   [ ON DELETE { NO ACTION | CASCADE | SET NULL | SET DEFAULT } ]
'   [ ON UPDATE { NO ACTION | CASCADE | SET NULL | SET DEFAULT } ]
'   [ NOT FOR REPLICATION ]
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

Public Class ForeignKeyColumn
    Private sName As String
    Private sLinked As String
    Private iSequence As Integer

#Region "Properties"
    Public Property Name() As String
        Get
            Name = sName
        End Get
        Set(ByVal nm As String)
            sName = nm
        End Set
    End Property

    Public Property Linked() As String
        Get
            Linked = sLinked
        End Get
        Set(ByVal lk As String)
            sLinked = lk
        End Set
    End Property

    Public ReadOnly Property Sequence() As Integer
        Get
            Sequence = iSequence
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub New(ByVal ColumnName As String, _
                   ByVal LinkedColumn As String, _
                   ByVal SequenceNo As Integer)
        sName = ColumnName
        sLinked = LinkedColumn
        iSequence = SequenceNo
    End Sub
#End Region

End Class

Public Class ForeignKeyColumns
    Inherits CollectionBase

#Region "Properties"
    Default Public Overloads ReadOnly Property Item(ByVal Index As Integer) As ForeignKeyColumn
        Get
            Return CType(List.Item(Index), ForeignKeyColumn)
        End Get
    End Property

    Default Public Overloads ReadOnly Property Item(ByVal Name As String) As ForeignKeyColumn
        Get
            For Each fkc As ForeignKeyColumn In Me
                If fkc.Name = Name Then
                    Return fkc
                End If
            Next
            Return Nothing
        End Get
    End Property
#End Region

#Region "Methods"
    Public Function Add(ByVal ColumnName As String, ByVal LinkedColumn As String, ByVal Sequence As Integer) As ForeignKeyColumn
        Dim fkc As New ForeignKeyColumn(ColumnName, LinkedColumn, Sequence)
        List.Add(fkc)
        Return fkc
    End Function
#End Region

End Class

Public Class ForeignKey
    Private sSchema As String
    Private qSchema As String
    Private sTable As String
    Private qTable As String
    Private sName As String = ""
    Private qName As String = ""
    Private sLinkedSchema As String
    Private sLinkedTable As String
    Private sMatch As String
    Private sDelete As String
    Private sUpdate As String
    Private bReplicated As Boolean = False

    Private sqllib As New sql
    Private cCols As New ForeignKeyColumns

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

    Public Property LinkedSchema() As String
        Get
            LinkedSchema = sLinkedSchema
        End Get
        Set(ByVal ls As String)
            sLinkedSchema = ls
        End Set
    End Property

    Public Property LinkedTable() As String
        Get
            LinkedTable = sLinkedTable
        End Get
        Set(ByVal lt As String)
            sLinkedTable = lt
        End Set
    End Property

    Public Property MatchOption() As String
        Get
            MatchOption = sMatch
        End Get
        Set(ByVal mt As String)
            sMatch = mt
        End Set
    End Property

    Public Property DeleteOption() As String
        Get
            DeleteOption = sDelete
        End Get
        Set(ByVal dlt As String)
            sDelete = dlt
        End Set
    End Property

    Public Property UpdateOption() As String
        Get
            UpdateOption = sUpdate
        End Get
        Set(ByVal uo As String)
            sUpdate = uo
        End Set
    End Property

    Public Property Replicated() As Boolean
        Get
            Replicated = bReplicated
        End Get
        Set(ByVal rp As Boolean)
            bReplicated = rp
        End Set
    End Property

    Public Property Columns() As ForeignKeyColumns
        Get
            Columns = cCols
        End Get
        Set(ByVal cc As ForeignKeyColumns)
            cCols = cc
        End Set
    End Property

    Public ReadOnly Property ForeignKeyShort() As String
        Get
            Dim ss As String = ""
            Dim st As String = ""
            Dim sOut As String

            sOut = "alter table " & qSchema & "." & qTable
            sOut &= " add constraint " & qName & vbCrLf
            sOut &= "foreign key ("
            For Each fkc As ForeignKeyColumn In cCols
                sOut &= st & sqllib.QuoteIdentifier(fkc.Name)
                ss &= st & sqllib.QuoteIdentifier(fkc.Linked)
                st = ","
            Next
            sOut &= ") references "
            sOut &= sqllib.QuoteIdentifier(sLinkedSchema) & "."
            sOut &= sqllib.QuoteIdentifier(sLinkedTable) & "("
            sOut &= ss & ")"
            If bReplicated Then
                sOut &= " not for replication"
            End If
            sOut &= vbCrLf

            st = ""
            If sDelete <> "NO ACTION" Then
                sOut &= "on delete " & sDelete
                st = " "
            End If

            If sUpdate <> "NO ACTION" Then
                sOut &= st & "on update " & sUpdate
                st = " "
            End If

            If st <> "" Then
                sOut &= vbCrLf
            End If

            ForeignKeyShort = sOut
        End Get
    End Property

    Public ReadOnly Property ForeignKeyText() As String
        Get
            Dim i As Integer = 0
            Dim ss As String = ""
            Dim sOut As String
            Dim st As String

            sOut = "declare @c1 integer, @c2 integer" & vbCrLf
            sOut &= vbCrLf
            sOut &= "if object_id('" & sName & "') is not null" & vbCrLf
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
            For Each fkc As ForeignKeyColumn In cCols
                If i = 0 Then
                    sOut &= "        select  " & fkc.Sequence & " keyno, '"
                    sOut &= fkc.Name & "' lkey, '"
                    sOut &= fkc.Linked & "' fkey" & vbCrLf
                    i = 1
                Else
                    sOut &= "        union select  " & fkc.Sequence
                    sOut &= ", '" & sqllib.QuoteIdentifier(fkc.Name)
                    sOut &= "', '" & sqllib.QuoteIdentifier(fkc.Linked) & "'" & vbCrLf
                End If
            Next
            sOut &= "    ) x" & vbCrLf
            sOut &= "    on      x.keyno = u1.ORDINAL_POSITION" & vbCrLf
            sOut &= "    and     x.lkey = u1.COLUMN_NAME" & vbCrLf
            sOut &= "    and     x.fkey = u2.COLUMN_NAME" & vbCrLf
            sOut &= "    where   c.CONSTRAINT_NAME = '" & sName & "'" & vbCrLf
            sOut &= "    and     c.DELETE_RULE = '" & sDelete & "'" & vbCrLf
            sOut &= "    and     c.UPDATE_RULE = '" & sUpdate & "'" & vbCrLf
            sOut &= "    and     objectproperty(object_id('" & sSchema & "." & sName & "'),'CnstIsNotRepl') = "
            If bReplicated Then
                sOut &= "1" & vbCrLf
            Else
                sOut &= "0" & vbCrLf
            End If
            sOut &= "    and     u1.TABLE_NAME = '" & sTable & "'" & vbCrLf
            sOut &= "    and     u2.TABLE_NAME = '" & sLinkedTable & "'" & vbCrLf
            sOut &= vbCrLf
            sOut &= "    if coalesce(@c1,0) <> coalesce(@c2,0) or coalesce(@c1,0) <> " & cCols.Count & vbCrLf
            sOut &= "    begin" & vbCrLf
            sOut &= "        print 'changing foreign key ''" & sName & "'''" & vbCrLf
            sOut &= "        alter table " & qSchema & "." & qTable & " drop constraint " & qName & vbCrLf
            sOut &= "    end" & vbCrLf
            sOut &= "end" & vbCrLf
            sOut &= vbCrLf
            sOut &= "if object_id('" & sName & "') is null" & vbCrLf
            sOut &= "begin" & vbCrLf
            sOut &= "    print 'creating foreign key ''" & sName & "'''" & vbCrLf
            sOut &= "    alter table " & qSchema & "." & qTable & " add constraint " & qName & vbCrLf
            sOut &= "    foreign key ("
            st = ""
            For Each fkc As ForeignKeyColumn In cCols
                sOut &= st & sqllib.QuoteIdentifier(fkc.Name)
                ss &= st & sqllib.QuoteIdentifier(fkc.Linked)
                st = ","
            Next
            sOut &= ") references "
            sOut &= sqllib.QuoteIdentifier(sLinkedSchema) & "."
            sOut &= sqllib.QuoteIdentifier(sLinkedTable) & "("
            sOut &= ss & ")"
            If bReplicated Then
                sOut &= " not for replication"
            End If
            sOut &= vbCrLf

            st = "    "
            If sDelete <> "NO ACTION" Then
                sOut &= st & "on delete " & sDelete
                st = " "
            End If

            If sUpdate <> "NO ACTION" Then
                sOut &= st & "on update " & sUpdate
                st = " "
            End If

            If st <> "" Then
                sOut &= vbCrLf
            End If
            sOut &= "end" & vbCrLf

            ForeignKeyText = sOut
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub New(ByVal Schema As String, ByVal Table As String, ByVal KeyName As String)
        Me.Name = KeyName
        sSchema = Schema
        qSchema = sqllib.QuoteIdentifier(sSchema)
        sTable = Table
        qTable = sqllib.QuoteIdentifier(sTable)
    End Sub

    Public Function ForeignKeyXML(ByVal sTab As String) As String
        Dim sOut As String

        sOut = sTab & "<foreignkey name='" & sName & "'"
        sOut &= " references='" & sLinkedTable & "'"
        If sDelete <> "NO ACTION" Then
            sOut &= " ondelete='" & sDelete & "'"
        End If
        If sUpdate <> "NO ACTION" Then
            sOut &= " onupdate='" & sUpdate & "'"
        End If
        If bReplicated Then
            sOut &= " replication='N'"
        End If
        sOut &= ">" & vbCrLf
        For Each fkc As ForeignKeyColumn In cCols
            sOut &= "  <column name='" & fkc.Name & "'"
            sOut &= " linksto='" & fkc.Linked & "' />" & vbCrLf
        Next
        sOut &= sTab & "</foreignkey>" & vbCrLf
        Return sOut
    End Function
#End Region

End Class

Public Class ForeignKeys
    Inherits CollectionBase

#Region "Properties"
    Default Public Overloads ReadOnly Property Item(ByVal KeyName As String) As ForeignKey
        Get
            For Each fk As ForeignKey In Me
                If fk.Name = KeyName Then
                    Return fk
                End If
            Next
            Return Nothing
        End Get
    End Property
#End Region

#Region "Methods"
    Public Function Add(ByVal forkey As ForeignKey) As Integer
        Return List.Add(forkey)
    End Function
#End Region
End Class

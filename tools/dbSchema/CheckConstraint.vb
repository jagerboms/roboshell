Option Explicit On
Option Strict On

' [ constraint CONSTRAINT_NAME ] check { [ not for replication ] ( EXPRESSION ) }

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

Public Class CheckConstraint
    Private sSchema As String
    Private qSchema As String
    Private sTable As String
    Private qTable As String
    Private sName As String = ""
    Private qName As String = ""
    Private sColumnName As String = ""
    Private sDefinition As String = ""
    Private bReplicated As Boolean = False
    Private bSystemName As Boolean = False
    Private bSystem As Boolean = False

    Private sqllib As New sql

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

    Public Property ColumnName() As String
        Get
            ColumnName = sColumnName
        End Get
        Set(ByVal cn As String)
            sColumnName = cn
        End Set
    End Property

    Public Property Definition() As String
        Get
            Definition = sDefinition
        End Get
        Set(ByVal df As String)
            sDefinition = df
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

    Public Property SystemName() As Boolean
        Get
            SystemName = bSystemName
        End Get
        Set(ByVal sn As Boolean)
            bSystemName = sn
        End Set
    End Property

    Public Property IsSystem() As Boolean
        Get
            IsSystem = bSystem
        End Get
        Set(ByVal sn As Boolean)
            bSystem = sn
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub New(ByVal Schema As String, ByVal Table As String, ByVal CheckName As String)
        Me.Name = CheckName
        sSchema = Schema
        qSchema = sqllib.QuoteIdentifier(sSchema)
        sTable = Table
        qTable = sqllib.QuoteIdentifier(sTable)
    End Sub

    Public Function Text(ByVal sTab As String, ByVal opt As ScriptOptions) As String
        Dim sOut As String = ""

        sOut = sTab & ","
        If opt.CheckShowName And Not bSystemName Then
            sOut &= "constraint " & qName & " "
        End If
        sOut &= "check"
        If bReplicated Then
            sOut &= " not for replication"
        End If
        sOut &= " (" & sqllib.CleanConstraint(sDefinition) & ")" & vbCrLf
        Return sOut
    End Function

    Public Function XMLText(ByVal sTab As String, ByVal opt As ScriptOptions) As String
        Dim sOut As String

        sOut = sTab & "<constraint type='check'"
        If opt.CheckShowName And Not bSystemName Then
            sOut &= " name='" & sName & "'"
        End If
        If bReplicated Then
            sOut &= " replication='N'"
        End If
        If sColumnName <> "" Then
            sOut &= " column='" & sColumnName & "'"
        End If
        sOut &= ">" & vbCrLf
        sOut &= sTab & "  <![CDATA["
        sOut &= sqllib.CleanConstraint(sDefinition)
        sOut &= "]]>" & vbCrLf
        sOut &= sTab & "</constraint>" & vbCrLf
        Return sOut
    End Function
#End Region

End Class

Public Class CheckConstraints
    Inherits CollectionBase

#Region "Properties"
    Default Public Overloads ReadOnly Property Item(ByVal CheckName As String) As CheckConstraint
        Get
            For Each cc As CheckConstraint In Me
                If cc.Name = CheckName Then
                    Return cc
                End If
            Next
            Return Nothing
        End Get
    End Property
#End Region

#Region "Methods"
    Public Function Add(ByVal chkcon As CheckConstraint) As Integer
        Return List.Add(chkcon)
    End Function

    Public Function XMLText(ByVal sTab As String, ByVal opt As ScriptOptions) As String
        Dim ss As String = ""
        Dim sOut As String = ""
        Dim cCC As CheckConstraint

        For Each cCC In Me
            ss &= cCC.XMLText(sTab & "  ", opt)
        Next
        If ss <> "" Then
            sOut &= sTab & "<constraints>" & vbCrLf
            sOut &= ss
            sOut &= sTab & "</constraints>" & vbCrLf
        End If
        Return sOut
    End Function
#End Region
End Class

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

Public Class FileGroup
    Private sqllib As New sql
    Private sDefaultGroup As String = "PRIMARY"
    Private sTableGroup As String = ""
    Private sTextGroup As String = ""
    Private sIndexGroup As String = ""
    Private sScheme As String = ""
    Private sSchemeColumn As String = ""

#Region "Properties"
    Public Property DefaultGroup() As String
        Get
            DefaultGroup = sDefaultGroup
        End Get
        Set(ByVal dfg As String)
            sDefaultGroup = dfg
        End Set
    End Property

    Public Property TableGroup() As String
        Get
            Dim s As String
            If sTableGroup = "" Then
                s = sDefaultGroup
            Else
                s = sTableGroup
            End If
            TableGroup = s
        End Get
        Set(ByVal tfg As String)
            sTableGroup = tfg
        End Set
    End Property

    Public Property TextGroup() As String
        Get
            Dim s As String
            If sTextGroup = "" Then
                If sTableGroup = "" Then
                    s = sDefaultGroup
                Else
                    s = sTableGroup
                End If
            Else
                s = sTextGroup
            End If
            TextGroup = s
        End Get
        Set(ByVal xfg As String)
            sTextGroup = xfg
        End Set
    End Property

    Public Property IndexGroup() As String
        Get
            Dim s As String
            If sIndexGroup = "" Then
                If sTableGroup = "" Then
                    s = sDefaultGroup
                Else
                    s = sTableGroup
                End If
            Else
                s = sIndexGroup
            End If
            IndexGroup = s
        End Get
        Set(ByVal ifg As String)
            sIndexGroup = ifg
        End Set
    End Property

    Public Property PartitionScheme() As String
        Get
            PartitionScheme = sScheme
        End Get
        Set(ByVal ps As String)
            sScheme = ps
        End Set
    End Property

    Public Property SchemeColumn() As String
        Get
            SchemeColumn = sSchemeColumn
        End Get
        Set(ByVal psc As String)
            sSchemeColumn = psc
        End Set
    End Property

    Public ReadOnly Property TableText() As String
        Get
            Dim s As String

            If sScheme <> "" Then
                s = " on " & sqllib.QuoteIdentifier(sScheme) & "("
                If sSchemeColumn = "" Then
                    s &= "??"
                Else
                    s &= sqllib.QuoteIdentifier(sSchemeColumn)
                End If
                s &= ")"
            Else
                s = GetTableText()
                If s <> "" Then
                    s = " on " & sqllib.QuoteIdentifier(s)
                End If
            End If
            TableText = s
        End Get
    End Property

    Public ReadOnly Property TableXML() As String
        Get
            Dim s As String

            If sScheme <> "" Then
                s = " partitionfunction='" & sScheme & "' partitioncolumn='"
                If sSchemeColumn = "" Then
                    s &= "??"
                Else
                    s &= sSchemeColumn
                End If
                s &= "'"
            Else
                s = GetTableText()
                If s <> "" Then
                    s = " filegroup='" & s & "'"
                End If
            End If
            TableXML = s
        End Get
    End Property

    Public ReadOnly Property TextText() As String
        Get
            Dim s As String

            If sTableGroup = "" Or sTableGroup = sDefaultGroup Then
                s = sDefaultGroup
            Else
                s = sTableGroup
            End If
            If sTextGroup = "" Or sTextGroup = s Then
                s = ""
            Else
                s = sTextGroup
            End If
            If s <> "" Then
                s = " textimage_on " & sqllib.QuoteIdentifier(s)
            End If
            TextText = s
        End Get
    End Property

    Public ReadOnly Property TextXML() As String
        Get
            Dim s As String

            If sTableGroup = "" Or sTableGroup = sDefaultGroup Then
                s = sDefaultGroup
            Else
                s = sTableGroup
            End If
            If sTextGroup = "" Or sTextGroup = s Then
                s = ""
            Else
                s = sTextGroup
            End If
            If s <> "" Then
                s = " textfilegroup='" & s & "'"
            End If
            TextXML = s
        End Get
    End Property

    Public ReadOnly Property IndexText() As String
        Get
            Dim s As String

            If sScheme <> "" Then
                s = " on " & sqllib.QuoteIdentifier(sScheme) & "("
                If sSchemeColumn = "" Then
                    s &= "??"
                Else
                    s &= sqllib.QuoteIdentifier(sSchemeColumn)
                End If
                s &= ")"
            Else
                s = GetIndexText()
                If s <> "" Then
                    s = " on " & sqllib.QuoteIdentifier(s)
                End If
            End If
            IndexText = s
        End Get
    End Property

    Public ReadOnly Property IndexXML() As String
        Get
            Dim s As String

            If sScheme <> "" Then
                s = " partitionfunction='" & sScheme & "' partitioncolumn='"
                If sSchemeColumn = "" Then
                    s &= "??"
                Else
                    s &= sSchemeColumn
                End If
                s &= "'"
            Else
                s = GetIndexText()
                If s <> "" Then
                    s = " filegroup='" & s & "'"
                End If
            End If
            IndexXML = s
        End Get
    End Property
#End Region

#Region "Methods"

#End Region

#Region "Library Routines"
    Private Function GetTableText() As String
        If sTableGroup = "" Or sTableGroup = sDefaultGroup Then
            GetTableText = ""
        Else
            GetTableText = sTableGroup
        End If
    End Function

    Private Function GetIndexText() As String
        Dim sI As String = Me.IndexGroup
        Dim sT As String = Me.TableGroup

        If sI = sT Then
            GetIndexText = ""
        Else
            GetIndexText = sI
        End If
    End Function
#End Region
End Class

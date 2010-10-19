Option Explicit On
Option Strict On

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

Public Class TablePermission
    Private sGrantee As String
    Private sGranted As String
    Private sPermissionName As String
    Private bDeny As Boolean = False
    Private bGrantOption As Boolean = False
    'Private sState As String
    Private aColumns As New ArrayList()

#Region "Properties"
    Public Property Grantee() As String
        Get
            Grantee = sGrantee
        End Get
        Set(ByVal gg As String)
            sGrantee = gg
        End Set
    End Property

    Public Property GrantedObject() As String
        Get
            GrantedObject = sGranted
        End Get
        Set(ByVal go As String)
            sGranted = go
        End Set
    End Property

    Public Property PermissionName() As String
        Get
            PermissionName = sPermissionName
        End Get
        Set(ByVal pn As String)
            sPermissionName = pn
        End Set
    End Property

    Public Property Deny() As Boolean
        Get
            Deny = bDeny
        End Get
        Set(ByVal dy As Boolean)
            bDeny = dy
        End Set
    End Property

    Public Property GrantOption() As Boolean
        Get
            GrantOption = bGrantOption
        End Get
        Set(ByVal op As Boolean)
            bGrantOption = op
        End Set
    End Property

    Public ReadOnly Property Columns() As ArrayList
        Get
            Columns = aColumns
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub New(ByVal pGrantee As String, _
            ByVal pGranted As String, _
            ByVal pPermissionName As String)

        sGrantee = pGrantee
        sGranted = pGranted
        sPermissionName = pPermissionName
    End Sub

    Public Sub AddColumn(ByVal ColumnName As String)
        aColumns.Add(ColumnName)
    End Sub
#End Region

End Class

Public Class TablePermissions
    Inherits CollectionBase

#Region "Properties"
    Default Public Overloads ReadOnly Property Item(ByVal Index As Integer) As TablePermission
        Get
            Return CType(List.Item(Index), TablePermission)
        End Get
    End Property

    Public ReadOnly Property Text() As String
        Get
            Dim sOut As String = ""
            Dim s As String
            Dim ss As String
            Dim sC As String

            For Each r As TablePermission In Me
                If r.Deny Then
                    sOut &= "deny "
                Else
                    sOut &= "grant "
                End If
                sOut &= LCase(r.PermissionName)

                sC = ""
                ss = ""
                For Each s In r.Columns
                    ss &= sC & s
                    sC = ", "
                Next
                If ss <> "" Then
                    sOut &= " (" & ss & ")"
                End If

                sOut &= " on " & r.GrantedObject
                sOut &= " to " & r.Grantee
                If r.GrantOption Then
                    sOut &= " with grant option"
                End If
                sOut &= vbCrLf
            Next
            Return sOut
        End Get
    End Property
#End Region

#Region "Methods"
    Public Function Add(ByVal Grantee As String, _
            ByVal Granted As String, _
            ByVal PermissionName As String) As TablePermission
        Dim fkc As New TablePermission(Grantee, Granted, PermissionName)
        List.Add(fkc)
        Return fkc
    End Function
#End Region

End Class

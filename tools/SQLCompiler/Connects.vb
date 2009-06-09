Option Explicit On
Option Strict On

Imports System.Data.SqlClient

Public Class Connect
    Private sName As String
    Private sConnectString As String
    Private sProvider As String
    Private sState As String = "U"

    Public Sub New(ByVal Name As String, ByVal ConnectString As String, ByVal Provider As String)
        sName = Name
        sConnectString = ConnectString
        sProvider = Provider
    End Sub

    Public ReadOnly Property Name() As String
        Get
            Name = sName
        End Get
    End Property

    Public ReadOnly Property ConnectString() As String
        Get
            ConnectString = sConnectString
        End Get
    End Property

    Public ReadOnly Property Provider() As String
        Get
            Provider = sProvider
        End Get
    End Property

    Public ReadOnly Property State() As String
        Get
            If sState = "U" Then CheckAccess()
            State = sState
        End Get
    End Property

    Private Sub CheckAccess()
        Dim psConn As SqlConnection
        Dim psAdapt As SqlDataAdapter
        Dim DS As DataSet

        Try
            If LCase(sProvider) <> "system.data.sqlclient" Then
                Return
            End If
            psConn = New SqlConnection(sConnectString)
            psConn.Open()
            psAdapt = New SqlDataAdapter("", psConn)
            psAdapt.SelectCommand.CommandText = "select is_member('db_owner')"
            DS = New DataSet
            psAdapt.Fill(DS, "data")

            If Not DS.Tables("data") Is Nothing Then
                If DS.Tables("data").Rows.Count > 0 Then
                    If DS.Tables("data").Rows(0).Item(0).ToString = "1" Then
                        sState = "OK"
                    End If
                End If
            End If
            psConn.Close()

        Catch ex As Exception
            Dim i As Integer = 9
        End Try
        If sState <> "OK" Then
            sState = "Error"
        End If
    End Sub
End Class

Public Class Connects

#Region "enumerator implementation"
    Implements IEnumerable
    Public Function GetEnumerator() As System.Collections.IEnumerator _
                    Implements System.Collections.IEnumerable.GetEnumerator
        Return New PropertyEnum(Keys, Values)
    End Function

    Public Class PropertyEnum
        Implements IEnumerable, IEnumerator
        Private Values As New Hashtable
        Dim Keys As ArrayList
        Private EnumeratorPosition As Integer = -1

        Public Sub New(ByVal aKeys As ArrayList, ByVal Hash As Hashtable)
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
                Return CType(Values.Item(Keys(EnumeratorPosition)), Connect)
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

    Private Values As New Hashtable
    Private Keys As New ArrayList

    Public ReadOnly Property List() As ArrayList
        Get
            List = Keys
        End Get
    End Property

    Public Function Add(ByVal Name As String, ByVal ConnectString As String, ByVal Provider As String) As Connect
        Dim parm As New Connect(Name, ConnectString, Provider)
        Values.Add(LCase(Name), parm)
        Keys.Add(LCase(Name))
        Return CType(Values.Item(LCase(Name)), Connect)
    End Function

    Public ReadOnly Property Item(ByVal Name As String) As Connect
        Get
            Try
                Return CType(Values.Item(LCase(Name)), Connect)
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property count() As Integer
        Get
            Return Values.Count
        End Get
    End Property
End Class

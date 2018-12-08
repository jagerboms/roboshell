Option Explicit On
Option Strict On

Imports System.Data.SqlClient

Public Class Connect
    Private sName As String
    Private sConnectString As String
    Private sProvider As String
    Private sState As String = "U"
    Private sErr As String = ""

    Public Sub New(ByVal Name As String, ByVal ConnectString As String, ByVal Provider As String)
        sName = Name
        sConnectString = ConnectString & ";Asynchronous Processing=true"
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

    Public ReadOnly Property ErrorText() As String
        Get
            ErrorText = sErr
        End Get
    End Property

    Private Sub CheckAccess()
        Dim slib As New sql()
        Dim s As String

        Try
            If LCase(sProvider) <> "system.data.sqlclient" Then
                Return
            End If

            slib.ConnectString = sConnectString
            s = slib.CheckAccess()

            If slib.CheckAccess() = "OK" Then
                sState = "OK"
            Else
                sErr &= "The user has insufficient privileges to run SQL compiler in the database." & vbCrLf
                sState = "Error"
            End If

        Catch ex As Exception
            sErr &= ex.Message & vbCrLf
            sState = "Error"
        End Try
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
